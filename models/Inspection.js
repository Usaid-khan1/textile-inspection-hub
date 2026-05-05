const mongoose = require('mongoose');

const CategorySchema = new mongoose.Schema({
  name: String,
  ok: Boolean,
  notok: Boolean,
  attachment: String
});

const DefectSchema = new mongoose.Schema({
  sno: Number,
  name: String,
  critical: String,
  major: String,
  minor: String
});

const CartonSchema = new mongoose.Schema({
  article: String,
  qty: String,
  ratio: String,
  selected: String,
  condition: String
});

const ImageCardSchema = new mongoose.Schema({
  caption: String,
  label: String,
  filename: String,       // stored filename on disk
  originalName: String,
  mimetype: String,
  url: String             // relative URL to serve the image
});

const InspectionSchema = new mongoose.Schema({
  // Header Info
  buyer: String,
  dpi: String,
  fri: String,
  poStyle: String,
  factory: String,
  inspDate: String,
  unit: String,
  reInspDate: String,

  // Article Info
  article: String,
  colour: String,
  orderQtyPcs: String,
  offeredQtyPcs: String,
  orderQtyCtns: String,
  offeredQtyCtns: String,
  unitsPerCarton: String,
  aqlCritical: String,
  aqlMajor: String,
  aqlMinor: String,
  reInspYes: Boolean,
  reInspNo: Boolean,
  approvedYes: Boolean,
  approvedNo: Boolean,
  inspectorName: String,
  inspLocation: String,

  // Result
  result: { type: String, enum: ['PASS', 'FAIL', 'PENDING', ''], default: '' },
  resultNote: String,

  // Sample
  samplePcs: String,
  selectedCartons: String,

  // AQL Allowed
  aqlAllowCritical: String,
  aqlAllowMajor: String,
  aqlAllowMinor: String,

  // Remarks
  inspectorRemarks: String,

  // Nested data
  categories: [CategorySchema],
  defects: [DefectSchema],
  cartons: [CartonSchema],
  imageCards: [ImageCardSchema],

  createdAt: { type: Date, default: Date.now },
  updatedAt: { type: Date, default: Date.now }
});

InspectionSchema.pre('save', function(next) {
  this.updatedAt = Date.now();
  next();
});

module.exports = mongoose.model('Inspection', InspectionSchema);
