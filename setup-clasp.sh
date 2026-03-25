# /bin/sh
# Script simples para preparar o setup local para usar o clasp

# Instala todos os programas necessários
brew install node npm
npm install -g @google/clasp
npm install --save @types/google-apps-script

# Solicita login
clasp login
