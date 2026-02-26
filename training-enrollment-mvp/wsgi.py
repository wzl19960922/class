import main

# Initialize shared resources when running under WSGI servers (e.g. gunicorn)
main.setup_logging()
main.initialize_database()

app = main.app
