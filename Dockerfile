FROM python:3.9.15-slim-bullseye

RUN python -m pip install --upgrade pip
RUN mkdir -p /notiz
COPY app /notiz/app/
RUN ls /notiz/app
RUN pip install -r /notiz/app/requirements.txt
ENTRYPOINT ['python', '/notiz/app/main.py']