# Utiliser une image de base avec Python
FROM python:3.11-slim

# Installer les dépendances système
RUN apt-get update && apt-get install -y \
    curl \
    git \
    unzip \
    xz-utils \
    zip \
    libglu1-mesa \
    && rm -rf /var/lib/apt/lists/*

# Installer Flutter
RUN git clone https://github.com/flutter/flutter.git -b stable /flutter
ENV PATH="/flutter/bin:${PATH}"

# Définir le répertoire de travail
WORKDIR /app

# Copier les fichiers de dépendances
COPY pubspec.yaml pubspec.lock ./
COPY backend/requirements.txt ./backend/

# Installer les dépendances Python
RUN pip install -r backend/requirements.txt

# Copier le reste du code
COPY . .

# Construire l'application Flutter (optionnel pour le backend)
RUN flutter pub get || echo "Flutter pub get failed, continuing..."

# Exposer le port
EXPOSE 5000

# Définir le répertoire de travail pour le backend
WORKDIR /app/backend

# Commande de démarrage
CMD ["python", "bon_a_envoye.py"] 