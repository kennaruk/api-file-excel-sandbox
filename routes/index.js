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

const mapExcelHeaders = {
	รหัสนักเรียน: "studentID",
	ชื่อจริง: "firstName",
	นามสกุล: "lastName",
	วันเกิด: "birthDate",
	รูปแบบการศึกษา: "educationType",
	ระดับการศึกษา: "educationLevel",
	ห้อง: "educationRoom"
};
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
function mapArrayToJsonByHeader(array, headers) {
	return new Promise((resolve, reject) => {
		let json = {};
		array.forEach((ele, i) => {
			const header = headers[i];
			const value = array[i];
			json[header] = value;
		});
		console.log("json", json);
		resolve(json);
	});
}
async function mapExcelDataToArrayJson(excelData) {
	try {
		let headers = {};
		/* Preparing headers to map */
		const excelHeaders = excelData.shift();

		for (let i = 0, { length } = excelHeaders; i < length; i++) {
			const headerValue = mapExcelHeaders[excelHeaders[i]];
			if (!headerValue) throw { message: "หัวตาราง excel ผิดรูปแบบ" };
			headers[i] = headerValue;
		}

		console.log("excelData", excelData);
		// console.log("excelData[0]", excelData[0]);
		// console.log("excelData[1]", excelData[1]);
		// console.log("excelData[2]", excelData[2]);

		/* Iterate and map each array to json */
		// for (let i = 0, { length } = excelData; i < length; i++) {
		// 	console.log("i", i);
		// 	const row = excelData[i];
		// 	// console.log("row", row);
		// 	excelData = await mapArrayToJsonByHeader(row, headers);
		// }

		return excelData;
	} catch (error) {
		throw error;
	}
}
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
		// const { rows, errors } = await readXlsxFile(savePath, { schema, sheet: 2 });
		const { rows, errors } = await readXlsxFile(savePath, { schema, sheet: 2 });
		fs.unlinkSync(savePath);

		console.log("rows::", rows);
		console.log("errors::", errors);

		if (errors.length !== 0) {
			res.status(400).send(errors);
			return;
		}
		/* */
		// console.log(await mapExcelDataToArrayJson(rows));
		// console.log("rows", rows);

		res.status(200).send(rows);
	} catch (error) {
		console.log(error);
		res.status(400).send(error);
	}
});
module.exports = router;
