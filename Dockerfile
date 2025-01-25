FROM python:3.11-bookworm

RUN pip install poetry==1.6.1

ENV POETRY_NO_INTERACTION=1 \
    POETRY_VIRTUALENVS_IN_PROJECT=1 \
    POETRY_VIRTUALENVS_CREATE=1 \
    POETRY_CACHE_DIR=/tmp/poetry_cache
WORKDIR /app

COPY pyproject.toml ./
RUN poetry install --no-dev --no-root && rm -rf $POETRY_CACHE_DIR
