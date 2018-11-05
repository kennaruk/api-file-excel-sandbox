var express = require("express");
var router = express.Router();
var fs = require("fs");
const fileUpload = require("express-fileupload");
const readXlsxFile = require("read-excel-file/node");
const fileExtension = require("file-extension");

router.use(fileUpload());
/* GET home page. */
router.get("/", function(req, res, next) {
	res.render("index", { title: "Express" });
});

router.get("/download", function(req, res, next) {
	const file = __dirname + "/../somefile.txt";
	res.download(file);
});

const excelHeaders = [
	"รหัสนักเรียน",
	"ชื่อจริง",
	"นามสกุล",
	"วันเกิด",
	"รูปแบบการศึกษา",
	"ระดับการศึกษา",
	"ห้อง"
];
router.post("/upload", async function(req, res) {
	try {
		const file = req.files.excel;
		const savePath = __dirname + "/../views/" + file.name;
		await file.mv(savePath);

		/* check extension */
		const extension = fileExtension(savePath);
		if (extension !== "xlsx") {
			res.status(400).send("not xlsx but", extension);
			return;
		}
		const rows = await readXlsxFile(savePath, { sheet: 2 });
		for (let i = 0; i < rows[0].length; i++) {
			if (rows[0][i] !== excelHeaders[i]) {
				console.log(rows[0][i], excelHeaders[i]);
			} else {
				console.log(excelHeaders[i], "true");
			}
		}
		console.log("rows", rows);
		fs.unlinkSync(savePath);

		res.status(200).json();
	} catch (error) {
		console.log(error);
	}
});
module.exports = router;
