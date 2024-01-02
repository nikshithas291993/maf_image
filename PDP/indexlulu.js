
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Function to read UPV numbers from an Excel sheet
function readUPVNumbersFromExcel(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0]; // Assuming UPV numbers are in the first sheet
  const worksheet = workbook.Sheets[sheetName];
  const range = xlsx.utils.decode_range(worksheet['!ref']);

  const upvNumbers = [];

  for (let rowNum = range.s.r + 1; rowNum <= range.e.r; rowNum++) {
    const cellAddress = { c: 1, r: rowNum }; // Assuming MAFAXRETAILVARIANTID is in the second column (B)
    const cellRef = xlsx.utils.encode_cell(cellAddress);
    const upvNumber = worksheet[cellRef].v;

    if (upvNumber) {
      upvNumbers.push(upvNumber);
    }
  }

  return upvNumbers;
}

// Function to make an API request and get the array of image URLs
async function getImageUrlsFromApi(apiEndpoint,upvNumber) {
  try {
    const response = await axios.get(apiEndpoint);
    const images = response.data.images;
    if (!images || !Array.isArray(images)) {
      throw new Error('Invalid or missing images in API response');
    }
    return images;
  } catch (error) {
    throw new Error(`Failed to get images from API: ${upvNumber}`);
  }
}
 
// Function to download an image from a URL and save it to a folder
async function downloadImage(url, folderPath, fileName) {
  try {
    const response = await axios({
      method: 'GET',
      url: url,
      responseType: 'stream',
    });
 
    const filePath = path.join(folderPath, fileName);
    const writer = fs.createWriteStream(filePath);
 
    response.data.pipe(writer);
 
    return new Promise((resolve, reject) => {
      writer.on('finish', () => resolve(filePath));
      writer.on('error', reject);
    });
  } catch (error) {
    throw new Error(`Failed to download image: ${error.message}`);
  }
}
 
const excelFilePath = './LULU_Product.xlsx';
const upvNumbers = readUPVNumbersFromExcel(excelFilePath);
const downloadFolder = './images_LULU'; // Change this to your desired folder path
 
 
Promise.all(
  upvNumbers.map(async (upvNumber) => {
    const apiEndpoint = `https://maf-ventures-prod.apigee.net/sap-proxy/rest/v2/lulu/products/${upvNumber}?fields=images(FULL)&lang=en&curr=AED`;
    let imageCounter = 1;
    try {
      const imageUrl = await getImageUrlsFromApi(apiEndpoint,upvNumber);
 
      for (let i = 0; i < imageUrl.length; i++) {
        const Url = imageUrl[i].url;
        const imageUrls = 'https://api.lululemon.me' + Url;
        const initialPart = imageUrl[i].code.match(/^([A-Z0-9_]+)_\d+/)[1];
 
        if (imageUrl[i].imageType === 'PRIMARY' && imageUrl[i].height === 1280) {
          const fileName = `${initialPart}_LULU_000_001.jpg`;
          await downloadImage(imageUrls, downloadFolder, fileName);
          console.log(`Image ${imageCounter} downloaded successfully.`);
          imageCounter++;
        }
        if (imageUrl[i].imageType === 'GALLERY' && imageUrl[i].height === 1280) {
          const fileName = `${initialPart}_LULU_000_${imageCounter.toString().padStart(3, '0')}.jpg`;
          await downloadImage(imageUrls, downloadFolder, fileName);
          console.log(`Image ${imageCounter} downloaded successfully.`);
          imageCounter++;
        }
      }
    } catch (error) {
      console.error(`Error for UPV number ${upvNumber}: ${error.message}`);
    }
  })
)
  .then(() => {
    console.log(`All images downloaded successfully`);
  })
  .catch((error) => {
    console.error(error.message);
  });
 