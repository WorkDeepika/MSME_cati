const mongoose = require("mongoose");

const projectSchema = new mongoose.Schema(
  {
    Projectname: { type: String, default: null ,required: true, unique: true},
    filteredCsv: { type: String, default: null },
    acceptedCsv: { type: String, default: null },
    startDate: { type: mongoose.Schema.Types.Mixed, default: null },
    endDate: { type: mongoose.Schema.Types.Mixed, default: null },
  },
  { timestamps: true }
);
const ProjectModel = mongoose.model("ProjectData", projectSchema);

module.exports = { ProjectModel};
