const {User}= require("../models/userModel")

const loginController= async (req, res)=>{
    const { email, password } = req.body;
  
      if (!email || !password) {
          return res.status(400).json({ message: 'Username and password are required' });
      }
  
      // Find the user by username
      const user = await User.findOne({ email : email });
      
      if (!user) {
          return res.status(401).json({ message: 'Invalid username or password' });
      }
  
      if (password != user.password) {
          return res.status(401).json({ message: 'Invalid username or password' });
      }
  
      return res.status(200).json({ 
          message: 'Login successful',
          user: {
              username: user.email
          },
          // token: token  // Uncomment if using JWT
      });
}
module.exports={
    loginController,
}