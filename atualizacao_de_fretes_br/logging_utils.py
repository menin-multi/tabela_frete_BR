import logging


def configure_logging(name: str) -> logging.Logger:
	logging.basicConfig(
		level=logging.INFO,
		format="%(asctime)s | %(message)s",
	)
	return logging.getLogger(name)