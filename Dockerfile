FROM python:3.11-alpine
WORKDIR /code
COPY requirements.txt /code/
COPY azure_devops_users_task.py /code/
RUN mkdir -p /code/logs /code/conf ; \
    pip install -r /code/requirements.txt 
CMD ["python", "azure_devops_users_task.py"]