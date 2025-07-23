# ✅ Step 1: Base image with Wine + Python + PyInstaller
FROM cdrx/pyinstaller-linux-windows:latest

# ✅ Step 2: Set working directory
WORKDIR /app

# ✅ Step 3: Copy your app code to the container
COPY . .

# ✅ Step 4: Install Python dependencies (optional)
# RUN pip install -r requirements.txt

# ✅ Step 5: Build the Windows EXE
RUN pyinstaller --onefile --windowed app.py

# ✅ Step 6: Output .exe to a shared volume
RUN mkdir -p /output && cp dist/app.exe /output/
