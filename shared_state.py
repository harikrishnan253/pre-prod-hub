from threading import Lock

WORD_LOCK = Lock()
download_tokens = {}
download_tokens_lock = Lock()
