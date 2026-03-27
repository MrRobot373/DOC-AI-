import multiprocessing

bind = "0.0.0.0:10000"
workers = 1  # Forced to 1 to allow the in-memory 'reviews' dictionary to work across all requests
threads = 8
worker_class = "gthread"
timeout = 300  # 5 minutes for long LLM document reviews
keepalive = 5
