const express = require('express');
const router = express.Router();
const ctrl = require('../controllers/inspectionController');
const upload = require('../middleware/upload');

// Inspections CRUD
router.get('/',              ctrl.getAllInspections);
router.get('/:id',           ctrl.getInspection);
router.post('/',             ctrl.createInspection);
router.put('/:id',           ctrl.updateInspection);
router.delete('/:id',        ctrl.deleteInspection);

// Image upload (single image)
router.post('/upload/image', upload.single('image'), ctrl.uploadImage);

// Excel download
router.get('/:id/excel',     ctrl.downloadExcel);

module.exports = router;
