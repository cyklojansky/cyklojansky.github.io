install:
	npm install

generate:
	python3 cyklojansky.py

dev:
	python3 -m http.server -d build/