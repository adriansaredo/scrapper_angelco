FROM python:3.8.4-slim-buster
COPY . app
WORKDIR /app

RUN pip install -r requirements.txt

CMD [ "main.py" ]

ENTRYPOINT ["python"]