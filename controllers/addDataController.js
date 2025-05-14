const {DataModel} = require('../models/dataModel');
// const {ProjectModel} = require('../models/projectModel');
const {User} = require('../models/userModel');
// const { Pool } = require('pg');
// const pool = new Pool({ connectionString: process.env.DATABASE_URL });
const putData = async (req, res) => {
    try {
        // Validate input data
        let jsonData = req.body;
        if (!jsonData || Object.keys(jsonData).length === 0) {
            return res.status(400).json({
                message: 'Request body is empty. Please provide valid data.',
            });
        }
        const result = await DataModel.create(jsonData);

        res.status(201).json({
            message: 'Data inserted successfully',
            result,
        });
    } catch (error) {
        // Handle Mongoose validation errors
        if (error.name === 'ValidationError') {
            return res.status(400).json({
                message: 'Validation error',
                error: error.message,
            });
        }

        // Handle unexpected errors
        res.status(500).json({
            message: 'Error inserting data',
            error: error.message,
        });
    }
};
const addUser = async (req, res) => {
    try {
        // Validate input data
        let jsonData = req.body;
        if (!jsonData || Object.keys(jsonData).length === 0) {
            return res.status(400).json({
                message: 'Request body is empty. Please provide valid data.',
            });
        }
        const result = await User.create(jsonData);

        res.status(201).json({
            message: 'Data inserted successfully',
            result,
        });
    } catch (error) {
        // Handle Mongoose validation errors
        if (error.name === 'ValidationError') {
            return res.status(400).json({
                message: 'Validation error',
                error: error.message,
            });
        }

        // Handle unexpected errors
        res.status(500).json({
            message: 'Error inserting data',
            error: error.message,
        });
    }
};
// const addUser= async (req, res) => {
//     const client = await pool.connect();
//     try {
//         const data = req.body;

//         // Mapping input fields to correct database columns
//         const name = data.name ? data.name.trim() : "Unknown";
//         const front_side_image = data.front_side_image ? data.front_side_image.trim() : "Unknown";
//         const back_side_image = data.back_side_image ? data.back_side_image.trim() : "Unknown";
//         const TeamHead = data.TeamHead ? data.TeamHead.trim() : "Unknown";
//         const city = data.city ? data.city.trim() : "Unknown";
//         const father_name = data.fatherName ? data.fatherName.trim() : "Unknown";
//         const gender = data.Gender ? data.Gender.trim() : "Other";
//         const dob = formatDate(data.dob);
//         const complete_address = data.address ? data.address.trim() : "Unknown";
//         const aadhar_no = data.AadharNO? data.AadharNO: "Unknown";
//         const pan_no = data.PanNo?data.PanNo: "Unknown";
//         const mobile_no = data.contact ? data.contact.trim() : "0000000000";
//         const email_id = data.Email ? data.Email.trim().toLowerCase() : null;
//         const education = data.Education ? data.Education.trim() : "Not Specified";
//         const created_at = new Date();
//         const updated_at = new Date(); // Auto-update timestamp

//         // Insert into database
//         const query = `
//             INSERT INTO users (name,front_side_image, back_side_image,TeamHead, city, father_name, gender, dob, complete_address, aadhar_no, pan_no, 
//                               mobile_no, email_id, education, created_at, updated_at) 
//             VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16)
//         `;

//         const result= await client.query(query, [
//             name,front_side_image, back_side_image, TeamHead, city, father_name, gender, dob, complete_address, aadhar_no, pan_no,
//             mobile_no, email_id, education, created_at, updated_at
//         ]);

//         res.status(201).json({ success: true, user: result });
//     } catch (error) {
//         console.error(`❌ Error inserting user:`, error);
//         res.status(500).json({ success: false, error: error.message });
//     } finally {
//         client.release();
//     }
// };

// const createOrUpdateProject = async (req, res) => {
//     try {
//       const { Projectname, ...projectData } = req.body;
  
//       let project = await ProjectModel.findOneAndUpdate(
//         { Projectname },
//         { $set: projectData },
//         { new: true, upsert: true } // upsert ensures a new document is created if name doesn't exist
//       );
  
//       return res.status(200).json({ success: true, projectName: project.Projectname, project });
//     } catch (error) {
//       console.error("Error handling project:", error);
//       return res.status(500).json({ success: false, message: "Internal Server Error" });
//     }
// };
// function formatDate(dateStr) {
//     if (!dateStr) return null;
//     if (dateStr.includes("/")) {
//         const [day, month, year] = dateStr.split("/");
//         return `${year}-${month}-${day}`;
//     }
//     return dateStr; // Already in YYYY-MM-DD format
// }

module.exports={ 
    putData,
    addUser,
    // createOrUpdateProject
}
