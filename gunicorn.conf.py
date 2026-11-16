# Gunicorn reads this file automatically from the project folder.
# Large protocol PDFs take minutes to read, so a request may run longer than the default 30 seconds.
timeout = 900
graceful_timeout = 900
