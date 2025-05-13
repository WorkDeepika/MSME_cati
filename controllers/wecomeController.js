
const welcome=(req, res)=>{
    res.send('Hello from deepika jain');
}
// const { S3Client, ListBucketsCommand } = require('@aws-sdk/client-s3');
// const welcome= async (req, res) => {
//     try {
//         const { S3Client } = require('@aws-sdk/client-s3');
//         const client = new S3Client({ region: 'eu-north-1' });
//         await client.send(new ListBucketsCommand({})); // Test request
//         res.status(200).send('Connected to S3 successfully');
//     } catch (err) {
//         console.error('S3 Connection Test Failed:', err);
//         res.status(500).send(err.message);
//     }
// };

module.exports={
    welcome
};