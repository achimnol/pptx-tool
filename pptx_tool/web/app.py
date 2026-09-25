import asyncio
import contextlib
import importlib.metadata
from collections.abc import AsyncIterator

import anyio.to_thread
from litestar import Litestar, MediaType, Router, get
from litestar.config.allowed_hosts import AllowedHostsConfig
from litestar.di import Provide
from litestar.openapi.config import OpenAPIConfig
from litestar.static_files import create_static_files_router
from litestar.types import ControllerRouterHandler

from .config import WebConfig
from .routes import route_handlers
from .storage import RequestStorage

_NOT_BUILT_HTML = """<!doctype html>
<html><head><meta charset="utf-8"><title>pptx-tool</title></head>
<body><p>The web UI is not built.
Run <code>pnpm -C frontend install &amp;&amp; pnpm -C frontend build</code> and restart the server.</p>
<p>The API documentation is available at <a href="/schema">/schema</a>.</p></body></html>
"""


@get("/", media_type=MediaType.HTML, include_in_schema=False, sync_to_thread=False)
def frontend_not_built() -> str:
    return _NOT_BUILT_HTML


def _frontend_handler(web_config: WebConfig) -> ControllerRouterHandler:
    static_dir = web_config.static_dir
    if static_dir is not None and (static_dir / "index.html").is_file():
        router: Router = create_static_files_router(path="/", directories=[static_dir], html_mode=True)
        return router
    return frontend_not_built


def _get_version() -> str:
    try:
        return importlib.metadata.version("pptx-tool")
    except importlib.metadata.PackageNotFoundError:
        return "0.0.0"


async def _expire_downloads_periodically(storage: RequestStorage) -> None:
    interval = min(max(storage.download_ttl / 2, 1.0), 60.0)
    while True:
        await asyncio.sleep(interval)
        await anyio.to_thread.run_sync(storage.expire_downloads)


def create_app(web_config: WebConfig | None = None) -> Litestar:
    config = web_config or WebConfig()
    storage = RequestStorage(config.tmp_dir, config.tmp_quota, config.download_ttl)

    @contextlib.asynccontextmanager
    async def expire_downloads(_: Litestar) -> AsyncIterator[None]:
        task = asyncio.create_task(_expire_downloads_periodically(storage))
        try:
            yield
        finally:
            task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await task

    def provide_web_config() -> WebConfig:
        return config

    def provide_storage() -> RequestStorage:
        return storage

    return Litestar(
        route_handlers=[*route_handlers, _frontend_handler(config)],
        dependencies={
            "web_config": Provide(provide_web_config, use_cache=True, sync_to_thread=False),
            "storage": Provide(provide_storage, use_cache=True, sync_to_thread=False),
        },
        on_startup=[storage.cleanup_stale],
        lifespan=[expire_downloads],
        # Leave some room for the multipart encoding overhead around the uploaded file.
        request_max_body_size=config.max_upload_size + 1024 * 1024,
        allowed_hosts=AllowedHostsConfig(allowed_hosts=list(config.allowed_hosts)) if config.allowed_hosts else None,
        openapi_config=OpenAPIConfig(title="pptx-tool", version=_get_version(), path="/schema"),
    )
