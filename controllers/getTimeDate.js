const moment = require("moment-timezone");
const getDateTime = async (req, res) =>{
    try {
        const now = moment().tz('Asia/Kolkata');
        const currentDate = now.format('DD-MM-YYYY'); // Extract date
        const currentTime = now.format('HH:mm:ss');   // Extract time

        res.status(200).json({ date: currentDate, time: currentTime });

    } catch (error) {
        console.error('Error fetching date and time:', error);
        res.status(500).json({ error: 'Internal Server Error', message: error.message });
    }

}
module.exports={
    getDateTime,
}