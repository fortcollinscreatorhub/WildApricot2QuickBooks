# setup
sudo apt-get install python3-venv
python3 -m venv venv
. venv/bin/activate
pip install Flask
pip install intuitlib
pip install intuit-oauth
pip install python-quickbooks
pip install pyopenssl

# run
. venv/bin/activate
export FLASK_APP=app.py
export FLASK_ENV=development
flask run --cert=adhoc
