require('dotenv').config();
const { Pool } = require('pg');
const xlsx = require('xlsx');
const {s3Client}= require("../services/aws/s3");
const { PutObjectCommand} = require("@aws-sdk/client-s3");
const pool = new Pool({ connectionString: process.env.DATABASE_URL });
const filterData= async (req, res) => {
    try {
      const { age, limit = 30 } = req.body; // Accept filters from frontend
  
      // Convert age filter to a valid SQL condition
      let whereClause = '1=1';
      if (age) {
        whereClause += ` AND dob < CURRENT_DATE - INTERVAL '${age} years'`;
      }
  
      // Fetch data from PostgreSQL
      const query = `SELECT name, father_name, gender, dob, mobile_no, email_id 
                     FROM users WHERE ${whereClause} 
                     ORDER BY created_at DESC LIMIT $1`;
  
      const { rows } = await pool.query(query, [limit]);
  
      // Create an XLSX workbook and worksheet
      const workbook = xlsx.utils.book_new();
      const worksheet = xlsx.utils.json_to_sheet(rows);
      xlsx.utils.book_append_sheet(workbook, worksheet, 'Users');
  
      // Write Excel file to buffer
      const excelBuffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  
      // Upload to S3 directly from buffer
      const uploadParams = {
        Bucket: process.env.BUCKET_NAME,
        Key: `exports/users_${Date.now()}.xlsx`,
        Body: excelBuffer,
        ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      };
  
    //   const uploadResult = await s3Client.upload(uploadParams).promise();
    const command = new PutObjectCommand(uploadParams);
    const uploadResult = await s3Client.send(command);
  
      // Return S3 File URL
      res.json({ success: true, 
        fileUrl: `https://${uploadParams.Bucket}.s3.eu-north-1.amazonaws.com/${uploadParams.Key}`,
    });
    } catch (error) {
      console.error(error);
      res.status(500).json({ success: false, error: error.message });
    }
}
module.exports={
    filterData,
}