# Stage 1: Build backend
FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Set the working directory
WORKDIR /app

# Install necessary system packages
RUN apt-get update && \
    apt-get install -y gcc default-libmysqlclient-dev pkg-config && \
    apt-get clean

# Copy the requirements file and install dependencies
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copy the entire project to the working directory
COPY . /app/

# Expose the port (Gunicorn will run on port 8000)
EXPOSE 8000

# Run the Gunicorn server
CMD ["gunicorn", "gcia.wsgi:application", "--bind", "0.0.0.0:8000", "--workers", "4", "--threads", "2"]
