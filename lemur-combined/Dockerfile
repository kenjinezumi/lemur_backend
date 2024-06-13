# Use the official Python image
FROM python:3.12-slim

# Set the working directory
WORKDIR /app

# Copy the requirements file
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the application code
COPY . .

# Ensure the PORT environment variable is available to the container
ENV PORT 8080

# Expose the port
EXPOSE 8080

# Command to run the application
ENTRYPOINT ["gunicorn", "--timeout", "1800", "-b", ":8080", "main:app"]
