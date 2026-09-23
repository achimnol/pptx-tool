"""Write the OpenAPI schema of the web API, used to generate the frontend API types."""

import json
import sys
from pathlib import Path

from .app import create_app
from .config import WebConfig


def main(argv: list[str] | None = None) -> None:
    args = sys.argv[1:] if argv is None else argv
    if len(args) != 1:
        sys.exit("usage: python -m pptx_tool.web.openapi OUTPUT_JSON")
    app = create_app(WebConfig(static_dir=None))
    schema = app.openapi_schema.to_schema()
    Path(args[0]).write_text(json.dumps(schema, indent=2, sort_keys=True, ensure_ascii=False) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
