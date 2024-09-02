
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const XLSX = require('xlsx');
const express = require("express");
const app = express();
const multer = require("multer");



app.use(express.json())
app.use(express.urlencoded({extended: true}))
app.use(express.static(path.join(__dirname, 'public')))     //static images, video can be used in 
app.set('view engine', 'ejs') //used to render ejs file as html
// Define the source and destination directories
const sourceDir = path.join(__dirname, 'uploadedImages');
const destDir = path.join(__dirname, 'downloadedImages');
app.listen(3000, () => {
  console.log("Server is running on 3000... ");
}); 




// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, './'); // Directory where files will be saved
    },
    filename: function (req, file, cb) {
        cb(null, file.originalname); // Use the original file name
    }
});

const upload = multer({ storage: storage });

// Function to download an image and save it to the specified path
const downloadImage = async (url, filePath) => {
    const writer = fs.createWriteStream(filePath);
    const response = await axios({
        url,
        method: 'GET',
        responseType: 'stream',
    });
    response.data.pipe(writer);
    return new Promise((resolve, reject) => {
        writer.on('finish', resolve);
        writer.on('error', reject);
    });
};

// Function to process the Excel file
const processExcel = async (excelFilePath, imagesFolder) => {
    const workbook = XLSX.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);

    if (!fs.existsSync(imagesFolder)) {
        fs.mkdirSync(imagesFolder);
    }

    for (const row of data) {
        const imageUrl = row['Image Path']; // Change 'Image Path' to the appropriate column name

        if (imageUrl) {
            const imageFileName = path.basename(imageUrl);
            const imageFilePath = path.join(imagesFolder, imageFileName);

            try {
                console.log(`Downloading ${imageUrl} to ${imageFilePath}`);
                await downloadImage(imageUrl, imageFilePath);
                console.log(`Successfully saved ${imageFileName}`);
            } catch (error) {
                console.error(`Failed to download ${imageUrl}:`, error);
            }
        }
    }
};

// Handle file upload and processing
app.post('/upload', upload.single('file'), async (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
    }

    const excelFilePath = path.join(__dirname, './', req.file.filename);
    const imagesFolder = path.join(__dirname, 'uploadedImages');

    try {
        await processExcel(excelFilePath, imagesFolder);
        res.json({ message: 'Excel file processed successfully' });
    } catch (error) {
        res.status(500).json({ error: 'Failed to process Excel file' });
    }
});

// Serve static files (like the HTML form) from the current directory
app.use(express.static(__dirname));










// DOWNLOAD UPLOADED IMAGE 
// Ensure the destination directory exists, create if it does not
function downloadUploadeImage (){
  if (!fs.existsSync(destDir)) {
    fs.mkdirSync(destDir, { recursive: true });
  }

  // Read the contents of the source directory
  fs.readdir(sourceDir, (err, files) => {
    if (err) {
      console.error("Error reading source directory:", err);
      return;
    }

    // Filter for image files (you can adjust this based on your needs)
    const imageFiles = files.filter((file) =>
      /\.(jpg|jpeg|png|gif)$/i.test(file)
    );

    imageFiles.forEach((file) => {
      const srcPath = path.join(sourceDir, file);
      const destPath = path.join(destDir, file);

      // Copy each image file
      fs.copyFile(srcPath, destPath, (err) => {
        if (err) {
          console.error(`Error copying file ${file}:`, err);
        } else {
          console.log(`Copied ${file} to ${destDir}`);
        }
      });
    });
  });
}

// ROUTES 
app.get("/", (req, res)=>{
    let a = "sak"
    res.render("index", {a:a})

})

app.post("/upload", (req, res)=>{
    processExcel(excelFilePath, imagesFolder);
    res.redirect("/")
})

app.get("/download", (req, res)=>{
downloadUploadeImage()
    res.redirect("/");

})