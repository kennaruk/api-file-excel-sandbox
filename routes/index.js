var express = require("express");
var router = express.Router();
var fs = require("fs");
const fileUpload = require("express-fileupload");
const readXlsxFile = require("read-excel-file/node");

router.use(fileUpload());
/* GET home page. */
router.get("/", function(req, res, next) {
	res.render("index", { title: "Express" });
});

router.get("/download", function(req, res, next) {
	const file = __dirname + "/../somefile.txt";
	res.download(file);
});

router.post("/upload", async function(req, res) {
	try {
		const file = req.files.excel;
		const savePath = __dirname + "/../views/" + file.name;
		await file.mv(savePath);

		const rows = await readXlsxFile(savePath);
		console.log("rows", rows);
		fs.unlinkSync(savePath);

		res.status(200).json();
	} catch (error) {
		console.log(error);
	}
});
module.exports = router;
