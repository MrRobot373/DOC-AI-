import multiprocessing

bind = "0.0.0.0:10000"
workers = multiprocessing.cpu_count() * 2 + 1
threads = 4
worker_class = "gthread"
timeout = 300  # 5 minutes for long LLM document reviews
keepalive = 5
