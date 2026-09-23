import logging
import threading

from pptx_tool.log import PACKAGE_LOGGER, capture_log

logger = logging.getLogger("pptx_tool.test")


def test_capture_log_collects_messages() -> None:
    with capture_log() as log:
        logger.info("hello %s", "world")
        logger.info("multi\nline")
    logger.info("after the context")
    assert log.lines == ["hello world", "multi\nline"]
    assert log.text == "hello world\nmulti\nline\n"


def test_capture_log_removes_its_handler() -> None:
    handlers = list(PACKAGE_LOGGER.handlers)
    with capture_log():
        assert len(PACKAGE_LOGGER.handlers) == len(handlers) + 1
    assert PACKAGE_LOGGER.handlers == handlers


def test_capture_log_is_isolated_per_thread() -> None:
    barrier = threading.Barrier(2)
    results: dict[str, list[str]] = {}

    def worker(name: str) -> None:
        with capture_log() as log:
            barrier.wait()
            for i in range(3):
                logger.info("%s-%d", name, i)
            barrier.wait()
        results[name] = log.lines

    threads = [threading.Thread(target=worker, args=(name,)) for name in ("a", "b")]
    for t in threads:
        t.start()
    for t in threads:
        t.join()
    assert results == {"a": ["a-0", "a-1", "a-2"], "b": ["b-0", "b-1", "b-2"]}
