const express = require('express'); 
const connectDB = require('./services/db/mongo');
const cors = require('cors');
const { createObjectCsvWriter } = require('csv-writer');
const path = require('path');
const fs = require('fs');
const dotenv = require('dotenv');
const routes= require('./Routes/main')
const createTable = require("./services/db/neon");
const app = express(); 
app.use(express.json()); 
app.use(cors({
    origin: "*",
    methods: ["POST", "GET", "PUT"],
    credentials: true
}));


dotenv.config();
connectDB();
createTable();
app.use('/', routes);

app.listen(3001, () => {
    console.log(`Server running on port 3000`);
});