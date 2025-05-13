const mongoose = require('mongoose');


const userSchema = new mongoose.Schema({
  email: { 
    type: mongoose.Schema.Types.Mixed , default: null
  },
  password: { type: mongoose.Schema.Types.String, required: true }
}, {timestamps: true});


const User = mongoose.model('User', userSchema);

module.exports = {
    User
};


