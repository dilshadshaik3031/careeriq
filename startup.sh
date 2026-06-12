#!/bin/bash
gunicorn --bind=0.0.0.0:8000 --timeout 600 --workers 2 --threads 4 app:app
Set-Content -Path startup.sh -Value "#!/bin/bash`ngunicorn --bind=0.0.0.0:80 --timeout 600 --workers 2 --threads 4 app:app"