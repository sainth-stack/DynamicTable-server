const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors'); // Import the cors middleware

const app = express();
const PORT = process.env.PORT || 5000;

// Use the provided MongoDB connection URL
mongoose.connect('mongodb+srv://eswarwork194:fH22ptwIfDPaTrEn@table-cluster.hyj6ymj.mongodb.net/mern_excel_upload', {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

const upload = multer({ dest: 'uploads/' });

// Enable CORS for all routes
app.use(cors());

// Create MongoDB schema
const dataSchema = new mongoose.Schema({
  BU: String,
  TASKFORCE: String,
  ZMHQ: String,
  RMHQ: String,
  'ABM HQ': String,
  'TM HQ': String,
  'TM ID': Number,
  'TM Name': String,
  Month: String, // Change the type to String
  DRCODE: Number,
  'MSL CODE': String,
  DRNAME: String,
  SPECIALTY: String,
  Spl: String,
  Class: String,
  'SKU CODE': Number,
  'SKU Name': String,
  BRAND: String,
  Qty: Number,
  Value: Number,
  'Segment Leadership': String,
  'Status 2k': String,
  'Status 5k': String
});


const DataModel = mongoose.model('Data', dataSchema);
function excelDateToFormattedDate(serial) {
  const date = new Date((serial - 25569) * 86400 * 1000);

  const months = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
  ];

  const month = months[date.getMonth()];
  const year = date.getFullYear().toString().substr(-2); // Get the last two digits of the year

  return `${month}-${year}`;
}


app.delete('/data/delete', async (req, res) => {
  try {
    // Use mongoose to remove all documents from the collection
    const result = await DataModel.deleteMany({});

    res.status(200).json({
      message: 'All data deleted successfully',
      deletedCount: result.deletedCount,
    });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const excelData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    // Convert Excel date serial numbers to JavaScript Date objects
    excelData.forEach((item) => {
      if (item.Month) {
        item.Month = excelDateToFormattedDate(item.Month);
      }
    });
    console.log(excelData)
    // Insert data into MongoDB
    await DataModel.insertMany(excelData);

    res.status(200).send('Data uploaded successfully');
  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});


app.get('/data', async (req, res) => {
  try {
    const data = await DataModel.find();
    res.json(data);
  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});

app.get('/Qty-data', async (req, res) => {
  try {
    // Extract Spl value from the request query
    const filters= req.query.items;
    // Fetch data from MongoDB based on Spl and other conditions if needed
    const data = await DataModel.find();

    const aggregateData = (data, header) => {
      const aggregatedData = data.reduce((result, record) => {
        // Generate a key based on the value of the specified header
        const key = record[header] || '';
    
        // Initialize count for the key if not present
        result[key] = (result[key] || 0) + 1;
    
        return result;
      }, {});
    
      return aggregatedData;
    };
    
    const aggregatedResults = {};
    filters?.forEach(filter => {
      const result = aggregateData(data || [], filter.split(','));
      aggregatedResults[filter] = result;
    });
    res.json({aggregatedResults});
  } catch (error) {
    console.error(error);
    res.status(500).send('Internal Server Error');
  }
});


app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
