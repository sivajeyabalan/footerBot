# Use a base image with LibreOffice pre-installed
FROM debian:bullseye

# Update package lists and install dependencies
RUN apt-get update && apt-get install -y libreoffice python3 python3-pip

# Set the working directory inside the container
WORKDIR /app

# Copy project files into the container
COPY . .

# Install required Python packages
RUN pip3 install --no-cache-dir -r requirements.txt

# Set the default command to run the bot
CMD ["python3", "bot.py"]
