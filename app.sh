# /bin/bash
. venv/bin/activate
export FLASK_APP=app.py
export FLASK_ENV=development
export PYTHONPATH=.:./etc
flask run --cert=adhoc
