# Use the official Python image as the base image
FROM python:3.12-slim

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file into the container
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the application code into the container
COPY . .

# Expose the port that Flask will run on
EXPOSE 5000

# Set the environment variables for Flask
ENV FLASK_APP=app.py
ENV FLASK_ENV=production

# Run the Flask app
CMD ["flask", "run", "--host=0.0.0.0"]
