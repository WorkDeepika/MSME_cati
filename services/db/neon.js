// require("dotenv").config();

// const http = require("http");
// const { neon } = require("@neondatabase/serverless");

// const sql = neon(process.env.DATABASE_URL);

// const requestHandler = async (req, res) => {
//   const result = await sql`SELECT version()`;
//   const { version } = result[0];
//   res.writeHead(200, { "Content-Type": "text/plain" });
//   res.end(version);
// };

// http.createServer(requestHandler).listen(3000, () => {
//   console.log("Server running at http://localhost:3000");
// });
require("dotenv").config();
const { neon } = require("@neondatabase/serverless");

const sql = neon(process.env.DATABASE_URL);

const createTable = async () => {
  try {
    await sql`
      CREATE TABLE IF NOT EXISTS users (
        id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
        name TEXT ,
        father_name TEXT ,
        gender TEXT CHECK (gender IN ('Male', 'Female', 'Other')) ,
        dob DATE ,
        complete_address TEXT ,
        aadhar_no TEXT UNIQUE ,
        pan_no TEXT UNIQUE ,
        TeamHead TEXT UNIQUE ,
        mobile_no TEXT ,
        email_id TEXT UNIQUE ,
        education TEXT ,
        front_side_image TEXT ,
        back_side_image TEXT ,
        city TEXT,
        created_at TIMESTAMP DEFAULT NOW(),
        updated_at TIMESTAMP DEFAULT NOW()
      );
    `;
    console.log("✅ Users table created or already exists");
  } catch (error) {
    console.error("❌ Error creating users table:", error);
  }
};

module.exports = createTable;
