FROM python:3.11-slim

WORKDIR /app

# LibreOffice + poppler render the re-skin preview (pptx -> pdf -> png). Without them
# the deck still rebrands and downloads; only the live image preview falls back.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress poppler-utils fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p output meetings staging build_context saved_templates static/renders

EXPOSE 5000

CMD ["gunicorn", "wsgi:app", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "120"]
