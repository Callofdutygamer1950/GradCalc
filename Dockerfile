# Use an official Python runtime as a parent image
FROM python:3.11

# Install system dependencies
RUN apt-get update && apt-get install -y tesseract-ocr libtesseract-dev

# Set working directory inside the container
WORKDIR /app

# Copy the application files into the container
COPY . /app

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose port 8080 for Render
EXPOSE 8080

# Start the Flask app with Gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "Att2GradeCalc:app"]
