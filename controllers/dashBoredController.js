require("dotenv").config();
const { Pool } = require("pg");
const pool = new Pool({ connectionString: process.env.DATABASE_URL });
const {ProjectModel} = require('../models/projectModel');
const genralInfo= async (req, res) => {
    const client = await pool.connect();
    try {
        const totalRecordsQuery = "SELECT COUNT(*) FROM users";
        const uniqueLocationsQuery = "SELECT COUNT(DISTINCT complete_address) FROM users";
        const uniqueCitiesQuery = "SELECT COUNT(DISTINCT city) FROM users";
        const cityCountsQuery = "SELECT city, COUNT(*) AS count FROM users GROUP BY city";

        const totalRecordsResult = await client.query(totalRecordsQuery);
        const uniqueLocationsResult = await client.query(uniqueLocationsQuery);
        const uniqueCitiesResult = await client.query(uniqueCitiesQuery);
        const cityCountsResult = await client.query(cityCountsQuery);

        const cityCounts = {};
        cityCountsResult.rows.forEach(row => {
            cityCounts[row.city] = parseInt(row.count, 10);
        });

        client.release();
        return res.json({
            total_records: parseInt(totalRecordsResult.rows[0].count, 10),
            unique_locations: parseInt(uniqueLocationsResult.rows[0].count, 10),
            unique_cities: parseInt(uniqueCitiesResult.rows[0].count, 10),
            city_counts: cityCounts
        });
    } catch (error) {
        console.error("❌ Error fetching stats:", error);
        client.release();
        return res.status(500).json({ error: "Internal Server Error" });
    } 
}
const getAllProjects = async (req, res) => {
    try {
        const projects = await ProjectModel.find();
        res.json({ success: true, data: projects });
    } catch (error) {
        console.error("Error fetching projects:", error);
        res.status(500).json({ success: false, message: "Error fetching projects" });
    }
};

module.exports={
    genralInfo,getAllProjects
}
