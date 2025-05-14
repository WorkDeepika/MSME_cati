const express = require('express');
const {welcome}=require("../controllers/wecomeController")
// const { getAudioUrlController }= require("../controllers/audioController")
// const { getImageUrl }= require("../controllers/addImageController")
const { loginController,  }= require("../controllers/loginController")
const {putData, addUser, }= require("../controllers/addDataController")
const {exportDataToXLSX }=  require("../controllers/getCsvController")
const multer = require('multer');
const path = require('path');
// const { getDateTime } = require('../controllers/getTimeDate');
// const { genralInfo,getAllProjects } = require('../controllers/dashBoredController');
// const { filterData } = require('../controllers/filterdataController');
// const { tempAddData} = require('../controllers/tempAdd')
// const { getInfo } = require('../controllers/dashboredController');

// Configure multer storage
const storage = multer.memoryStorage(); // Store files in memory
const upload = multer({ storage });

const router = express.Router();

router.get('/', welcome)
router.post('/add-data',putData);
router.post('/login', loginController);
router.post('/add-user',addUser);
// router.post('/get-xlsx',filterData);
// router.post('/addOrUpdateProject',createOrUpdateProject);
router.get('/download-xlsx',exportDataToXLSX)
// router.get('/getCurrentTimeDate', getDateTime);
// router.post('/add-image', upload.single('file'), getImageUrl);
// router.get('/genralInfo', genralInfo)
// router.get('/getProject', getAllProjects)

module.exports = router;