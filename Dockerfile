FROM python:3.9-slim

# Install Poppler (Required for PDF handling)
RUN apt-get update && apt-get install -y \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy your code
COPY . .

# Expose Gradio port
EXPOSE 7860

# Run the app
CMD ["python", "app.py"]
