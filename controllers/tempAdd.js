const { TempModel } = require("../models/tempModel");

const tempAddData = async (req, res) => {
    try {
        // Validate input data
        let jsonData = req.body;
        if (!jsonData || Object.keys(jsonData).length === 0) {
            return res.status(400).json({
                message: 'Request body is empty. Please provide valid data.',
            });
        }
        const result = await TempModel.create(jsonData);

        return res.status(201).json({
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
        return res.status(500).json({
            message: 'Error inserting data',
            error: error.message,
        });
    }
};
module.exports={
    tempAddData
}