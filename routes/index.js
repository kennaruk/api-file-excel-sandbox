var express = require("express");
var router = express.Router();
var fs = require("fs");
const fileUpload = require("express-fileupload");
const readXlsxFile = require("read-excel-file/node");
const fileExtension = require("file-extension");
var json2xls = require("json2xls");

router.use(fileUpload());
/* GET home page. */
router.get("/", function(req, res, next) {
	res.render("index", { title: "Express" });
});
const output = [
	{
		studentID: 45663,
		firstName: "สุรศักดิ์",
		lastName: "ชาติดำรงค์",
		birthDate: "2001-12-06T12:00:00.000Z",
		educationType: "มัธยม",
		educationLevel: 6,
		educationRoom: 4
	},
	{
		studentID: 306,
		firstName: "เพชรน้ำหนึ่ง",
		lastName: "บุญประเสริฐ",
		birthDate: "2008-02-22T12:00:00.000Z",
		educationType: "ประถม",
		educationLevel: 4,
		educationRoom: 10
	},
	{
		studentID: 11235,
		firstName: "เกียรติขจร",
		lastName: "เหล่าวินัย",
		birthDate: "2002-03-27T12:00:00.000Z",
		educationType: "ปวช",
		educationLevel: 2,
		educationRoom: "-"
	}
];

router.get("/download", function(req, res, next) {
	const file = __dirname + "/../somefile.txt";
	var xls = json2xls(output);
	const exportName = `${new Date().getTime()}-download.xlsx`;
	const exportPath = `${__dirname}/../seed_files/${exportName}`;
	console.log("exportPath:", exportPath);

	fs.writeFileSync(exportPath, xls, "binary");
	res.download(file);
	fs.unlinkSync(exportPath);
});

const schema = {
	รหัสนักเรียน: {
		prop: "studentID",
		type: String
	},
	ชื่อจริง: {
		prop: "firstName",
		type: String
	},
	นามสกุล: {
		prop: "lastName",
		type: String
	},
	วันเกิด: {
		prop: "birthDate",
		type: Date
	},
	รูปแบบการศึกษา: {
		prop: "educationType",
		type: String
	},
	ระดับการศึกษา: {
		prop: "educationLevel",
		type: String
	},
	ห้อง: {
		prop: "educationRoom",
		type: String
	}
};
router.post("/upload", async function(req, res) {
	try {
		const file = req.files.excel;
		/* check extension */
		const extension = fileExtension(file.name);
		if (extension !== "xlsx") {
			res.status(400).send("not xlsx but", extension);
			return;
		}

		const saveName = `${new Date().getTime()}-${file.name}`;
		const savePath = `${__dirname}/../seed_files/${saveName}`;
		console.log("SavePath:", savePath);
		await file.mv(savePath);

		const { rows, errors } = await readXlsxFile(savePath, { schema, sheet: 2 });
		fs.unlinkSync(savePath);

		if (errors.length !== 0) {
			res.status(400).send(errors);
			return;
		}

		res.status(200).send(rows);
	} catch (error) {
		console.log(error);
		res.status(400).send(error);
	}
});
module.exports = router;
