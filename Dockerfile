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

RUN mkdir -p output meetings staging build_context static/renders

# The review page's slide previews are committed (static/previews/, built by
# scripts/prerender_master.py). This is a safety net: if the master deck changed and
# nobody re-ran the script, regenerate here, where LibreOffice IS present. A no-op --
# and instant -- when the committed previews already match the deck's content hash.
RUN python scripts/prerender_master.py || true

EXPOSE 5000

CMD ["gunicorn", "wsgi:app", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "120"]
