const { DataModel } = require('../models/dataModel');
const XLSX = require("xlsx");
const {s3Client}= require("../services/aws/s3");
const { PutObjectCommand} = require("@aws-sdk/client-s3");
const path = require('path');

//write bucket name
require("dotenv").config();

const exportDataToXLSX = async (req, res) => {
    try {
        // Step 1: Fetch data from MongoDB
        const data = await DataModel.find().lean();
        if (!data.length) {
            return res.status(404).json({ message: 'No data found' });
        }

        // Define the expected fields
        const expectedFields = [
            "_id",
            "email",
            "startDate",
            "startTime",
            "backST",
            "backSD",
            "address",
            "interviewerName",
            "interviewerId",
            "city",
            "language",
            "QS",
            "Q1",
            "Q2a",
            "Q2b",
            "Q2b_other",
            "Q3",
            "Q4",
            "Q5",
            "Q6",
            "Q7",
            "Q8",
            "Q9",
            "Q9_other",
            "Q10",
            "Q11",
            "Q12",
            "TC_mode",
            "TCQ1_1_Ans",
            "TCQ1_2_Ans",
            "TCQ1_3_Ans",
            "TCQ1_4_Ans",
            "TCQ1_5_Ans",
            "TCQ1_total",
            "TCQ2",
            "TCQ3_1_Ans",
            "TCQ3_1_Ans2",
            "TCQ3_2_Ans",
            "TCQ3_3_Ans",
            "TCQ3_4_Ans",
            "TCQ3_3_Ans2",
            "TCQ3_2_Ans2",
            "TCQ3_4_Ans2",
            "TCQ4_1_Ans",
            "TCQ4_2_Ans",
            "TCQ4_1_Ans2",
            "TCQ4_2_Ans2",
            "TCQ4_3_Ans",
            "TCQ4_3_Ans2",
            "SCQ1",
            "SCQPOE_1_Ans",
            "SCQPOE_2_Ans",
            "SCQPOE_3_Ans",
            "SCQPOE_1_cost",
            "SCQPOE_2_cost",
            "SCQPOE_3_cost",
            "SCQPOE_1_main_cost",
            "SCQPOE_2_main_cost",
            "SCQPOE_3_main_cost",
            "SCQTE_1_Ans",
            "SCQTE_2_Ans",
            "SCQTE_3_Ans",
            "SCQTE_1_cost",
            "SCQTE_2_cost",
            "SCQTE_3_cost",
            "SCQTE_1_reason",
            "SCQTE_2_reason",
            "SCQTE_3_reason",
            "SCQRR_1_Ans",
            "SCQRR_2_Ans",
            "SCQRR_3_Ans",
            "SCQRR_1_cost",
            "SCQRR_2_cost",
            "SCQRR_3_cost",
            "SCQRR_1_reason",
            "SCQRR_2_reason",
            "SCQRR_3_reason",
            "SCQMR_1_Ans",
            "SCQMR_2_Ans",
            "SCQMR_3_Ans",
            "SCQMR_1_cost",
            "SCQMR_2_cost",
            "SCQMR_3_cost",
            "SCQMR_1_reason",
            "SCQMR_2_reason",
            "SCQMR_3_reason",
            "endTime",
            "endDate",
            "duration",
            "backET",
            "backED"
          ];
          

        const formatDate = (dateString) => {
            if (!dateString || typeof dateString !== "string") {
                console.error("Invalid date input:", dateString);
                return null; // Return null or a default value
            }
        
            let delimiter = dateString.includes("/") ? "/" : dateString.includes("-") ? "-" : null;
        
            if (!delimiter) {
                console.error("Unexpected date format:", dateString);
                return null; // Handle unexpected formats
            }
        
            const parts = dateString.split(delimiter);
            
            if (parts.length !== 3) {
                console.error("Invalid date structure:", dateString);
                return null; // Ensure we get exactly 3 parts
            }
        
            let [part1, part2, year] = parts.map(part => part.trim());
        
            if (!year || isNaN(year)) {
                console.error("Invalid year in date:", dateString);
                return null;
            }
        
            // Determine if format is MM-DD-YYYY or DD-MM-YYYY
            let month, day;
            if (parseInt(part1) > 12) { 
                // If first part is greater than 12, it's likely DD-MM-YYYY
                day = part1;
                month = part2;
            } else {
                // Otherwise, assume it's MM-DD-YYYY
                month = part1;
                day = part2;
            }
        
            if (!month || !day || isNaN(month) || isNaN(day)) {
                console.error("Invalid day/month in date:", { month, day, year });
                return null;
            }
        
            return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year}`;
        };
        

        // Step 2: Format the data and enforce the correct field order
        const formattedData = data.map(row => {
            const orderedRow = {};
            expectedFields.forEach(field => {
                if (field === "Date" ) {
                    orderedRow[field] = formatDate(row[field]);
                } else if (field === "_id") {
                    // Convert _id to a string
                    orderedRow[field] = row[field] ? row[field].toString() : null;
                }else if (Array.isArray(row[field])) {
                    // Convert arrays to a comma-separated string
                    orderedRow[field] = row[field].join(", ");
                } else {
                    orderedRow[field] = row.hasOwnProperty(field) ? row[field] : null;
                }
            });
            return orderedRow;
        });
        // Step 1: Create a worksheet with formatted data
        const worksheet = XLSX.utils.json_to_sheet(formattedData);
        // Step 3: Add `expectedFields` starting at `AY2`
        XLSX.utils.sheet_add_aoa(worksheet, [expectedFields], { origin: "A1" });

        // Step 4: Create a workbook and append the worksheet
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');


        // Step 4: Convert the workbook to a buffer
        const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

        const date = new Date();
        //console.log(date)
        const optionsDate = {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        timeZone: 'Asia/Kolkata',
        };

        const optionsTime = {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false, // Use 24-hour format
        timeZone: 'Asia/Kolkata',
        };

        // Format the date as YYYY-MM-DD
        const formattedDate = new Intl.DateTimeFormat('en-US', optionsDate)
        .format(date)
        .replace(/\//g, '-');

        // Format the time as HH-MM-SS
        const formattedTime = new Intl.DateTimeFormat('en-US', optionsTime)
        .format(date)
        .replace(/:/g, '-');
        
          // Combine date and time for unique naming
          const fileName = `msmi-cati-${formattedDate}-${formattedTime}.xlsx`;
        const uploadParams = {
            Bucket: process.env.BUCKET_NAME, // Replace with your bucket name
            Key: `xlsx/${fileName}`,
            Body: excelBuffer,
            ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        };

        try {
            const command = new PutObjectCommand(uploadParams);
            const result = await s3Client.send(command);
            console.log('Excel file uploaded to S3:', result);

            return res.status(200).json({
                message: 'Excel file uploaded successfully to S3',
                downloadUrl: `https://${uploadParams.Bucket}.s3.us-east-1.amazonaws.com/xlsx/${fileName}`,
            });
        } catch (err) {
            console.error('Error uploading Excel file to S3:', err);
            return res.status(500).json({
                message: 'Error uploading Excel file to S3',
                error: err.message,
            });
        }
    } catch (error) {
        console.error('Error exporting data to Excel:', error);
        return res.status(500).json({ message: 'An error occurred while exporting data to Excel' });
    }
};

// const exportDataToXLSX = async (req, res) => {
//     try {
//         // Step 1: Fetch data from MongoDB
//         const data = await DataModel.find().lean();
//         if (!data.length) {
//             return res.status(404).json({ message: 'No data found' });
//         }

//         const expectedFields = [
//                         "_id",
//                         "email",
//                         "startDate",
//                         "startTime",
//                         "backST",
//                         "backSD",
//                         "address",
//                         "interviewerName",
//                         "interviewerId",
//                         "city",
//                         "language",
//                         "QS",
//                         "Q1",
//                         "Q2a",
//                         "Q2b",
//                         "Q2b_other",
//                         "Q3",
//                         "Q4",
//                         "Q5",
//                         "Q6",
//                         "Q7",
//                         "Q8",
//                         "Q9",
//                         "Q9_other",
//                         "Q10",
//                         "Q11",
//                         "Q12",
//                         "TC_mode",
//                         "TCQ1_1_Ans",
//                         "TCQ1_2_Ans",
//                         "TCQ1_3_Ans",
//                         "TCQ1_4_Ans",
//                         "TCQ1_5_Ans",
//                         "TCQ1_total",
//                         "TCQ2",
//                         "TCQ3_1_Ans",
//                         "TCQ3_1_Ans2",
//                         "TCQ3_2_Ans",
//                         "TCQ3_3_Ans",
//                         "TCQ3_4_Ans",
//                         "TCQ3_3_Ans2",
//                         "TCQ3_2_Ans2",
//                         "TCQ3_4_Ans2",
//                         "TCQ4_1_Ans",
//                         "TCQ4_2_Ans",
//                         "TCQ4_1_Ans2",
//                         "TCQ4_2_Ans2",
//                         "TCQ4_3_Ans",
//                         "TCQ4_3_Ans2",
//                         "SCQ1",
//                         "SCQPOE_1_Ans",
//                         "SCQPOE_2_Ans",
//                         "SCQPOE_3_Ans",
//                         "SCQPOE_1_cost",
//                         "SCQPOE_2_cost",
//                         "SCQPOE_3_cost",
//                         "SCQPOE_1_main_cost",
//                         "SCQPOE_2_main_cost",
//                         "SCQPOE_3_main_cost",
//                         "SCQTE_1_Ans",
//                         "SCQTE_2_Ans",
//                         "SCQTE_3_Ans",
//                         "SCQTE_1_cost",
//                         "SCQTE_2_cost",
//                         "SCQTE_3_cost",
//                         "SCQTE_1_reason",
//                         "SCQTE_2_reason",
//                         "SCQTE_3_reason",
//                         "SCQRR_1_Ans",
//                         "SCQRR_2_Ans",
//                         "SCQRR_3_Ans",
//                         "SCQRR_1_cost",
//                         "SCQRR_2_cost",
//                         "SCQRR_3_cost",
//                         "SCQRR_1_reason",
//                         "SCQRR_2_reason",
//                         "SCQRR_3_reason",
//                         "SCQMR_1_Ans",
//                         "SCQMR_2_Ans",
//                         "SCQMR_3_Ans",
//                         "SCQMR_1_cost",
//                         "SCQMR_2_cost",
//                         "SCQMR_3_cost",
//                         "SCQMR_1_reason",
//                         "SCQMR_2_reason",
//                         "SCQMR_3_reason",
//                         "endTime",
//                         "endDate",
//                         "duration",
//                         "backET",
//                         "backED"
//                       ];

//         // Step 2: Format the data
//         const formattedData = data.map(row => {
//             const orderedRow = {};
//             expectedFields.forEach(field => {
//                 orderedRow[field] = row[field] ?? null;
//             });
//             return orderedRow;
//         });

//         // Step 3: Create Excel
//         const worksheet = XLSX.utils.json_to_sheet(formattedData);
//         XLSX.utils.sheet_add_aoa(worksheet, [expectedFields], { origin: "A1" });
//         const workbook = XLSX.utils.book_new();
//         XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
//         const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

//         // Generate filename
//         const date = new Date();
//         const formattedDate = date.toISOString().split('T')[0];
//         const formattedTime = date.toTimeString().split(' ')[0].replace(/:/g, '-');
//         const fileName = `msme-${formattedDate}-${formattedTime}.xlsx`;

//         const uploadParams = {
//             Bucket: process.env.BUCKET_NAME,
//             Key: `xlsx/${fileName}`,
//             Body: excelBuffer,
//             ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
//         };

//         // Step 4: Upload to S3
//         const command = new PutObjectCommand(uploadParams);
//         await s3Client.send(command);
//         console.log('Excel file uploaded to S3.');

//         // Step 5: Generate download URL (Region-Specific)
//         const downloadUrl = `https://${uploadParams.Bucket}.s3.us-east-1.amazonaws.com/xlsx/${fileName}`;

//         return res.status(200).json({
//             message: 'Excel file uploaded successfully to S3',
//             downloadUrl: downloadUrl,
//         });
//     } catch (error) {
//         console.error('Error exporting data to Excel:', error);
//         return res.status(500).json({ 
//             message: 'An error occurred while exporting data to Excel',
//             error: error.message,
//         });
//     }
// };

module.exports = { 
    exportDataToXLSX
};
