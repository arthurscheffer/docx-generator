const express = require('express');
const multer = require('multer');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
const upload = multer(); // Pour traiter les formulaires de type multipart/form-data

// Route GET pour vérifier que le service fonctionne
app.get('/', (req, res) => {
    res.send('Le service est en marche !');
  });
  
// Route POST pour générer le document
app.post('/generate', upload.single('template'), (req, res) => {
  try {
    // Récupérer le contenu du template (.docx) depuis le champ 'template'
    const content = req.file.buffer;

    // Charger le fichier en tant qu'archive zip
    const zip = new PizZip(content);
    // Créer l'instance de docxtemplater
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Les données de remplacement sont envoyées dans le champ 'data'
    // Elles doivent être au format JSON (sous forme de chaîne)
    const data = JSON.parse(req.body.data);

    // Remplacer les placeholders par les valeurs fournies
    doc.render(data);

    // Générer le document final au format Node Buffer
    const buf = doc.getZip().generate({ type: 'nodebuffer' });

    // Définir les headers pour indiquer qu'il s'agit d'un fichier Word
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': 'attachment; filename="generated.docx"',
    });

    // Envoyer le document généré
    res.send(buf);
  } catch (error) {
    console.error('Erreur lors de la génération du document :', error);
    res.status(500).send({ error: error.message });
  }
});

// Lancer le service sur le port 3000 (modifiable si besoin)
app.listen(3000, () => console.log('Service de génération lancé sur le port 3000'));
