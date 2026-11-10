FROM python:3.13-slim

RUN apt-get update \
 && apt-get install -y --no-install-recommends tesseract-ocr tesseract-ocr-eng libglib2.0-0 libgl1 \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV PORT=10000 PYTHONUNBUFFERED=1
# One process: conversion jobs are tracked in memory. Threads serve requests.
CMD gunicorn app:app --workers 1 --threads 8 --timeout 600 --bind 0.0.0.0:$PORT
