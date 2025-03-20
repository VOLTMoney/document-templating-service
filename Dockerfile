FROM python:3.8.13-slim-buster

# Set working directory
WORKDIR /code

# Install dependencies & clean up cache
RUN apt-get update && apt-get install -y logrotate && rm -rf /var/lib/apt/lists/*

# Create a virtual environment
RUN python -m venv venv
ENV PATH="venv/bin:$PATH"

# Set environment variable for Gotenberg
ENV GOTENBERG_API_URL=http://gotenberg:3000

# Copy and install dependencies
COPY ./requirements.txt /code/requirements.txt
RUN pip install --no-warn-script-location --no-cache-dir --upgrade -r /code/requirements.txt

# Copy application files
COPY . /code

# Set up log rotation
RUN echo "/code/temp/*.log { \n\
        daily \n\
        rotate 7 \n\
        compress \n\
        missingok \n\
        notifempty \n\
    }" > /etc/logrotate.d/app_logs

# Automatically clean up old temp files on container start
RUN echo '#!/bin/sh\nrm -rf /code/temp/*' > /cleanup.sh && chmod +x /cleanup.sh

# Expose port
EXPOSE 4532

# Start the application
CMD ["sh", "-c", "/cleanup.sh && exec venv/bin/uvicorn main:app --host 0.0.0.0 --port 4532"]
