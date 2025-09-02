install:
	npm install
	pip install -r requirements.txt

generate:
	python3 cyklojansky.py

dev:
	python3 -m http.server -d build/