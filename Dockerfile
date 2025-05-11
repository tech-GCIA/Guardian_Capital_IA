# Stage 1: Build backend
FROM python:3.11-slim

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Set the working directory
WORKDIR /app

# Copy the requirements file and install dependencies
COPY ./GUARCIAN_CAPITAL_IA/requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copy the entire project to the working directory
COPY ./GUARCIAN_CAPITAL_IA /app/

# Expose the port (Gunicorn will run on port 8000)
EXPOSE 8000

# Run the Gunicorn server
CMD ["gunicorn", "gcia.wsgi:application", "--bind", "0.0.0.0:8000", "--workers", "4", "--threads", "2"]
