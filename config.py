import os

# File paths configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
BACKUP_FOLDER = 'backups'
DEFAULT_SOURCE_FILE = os.path.join(UPLOAD_FOLDER, 'base_sheet.xlsx')
DEFAULT_NEW_DATA_FILE = os.path.join(UPLOAD_FOLDER, 'new_data.xlsx')

# Ensure required directories exist
for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER, BACKUP_FOLDER]:
    os.makedirs(folder, exist_ok=True)

# File configuration
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
