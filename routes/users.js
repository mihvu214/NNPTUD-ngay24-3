var express = require("express");
var router = express.Router();
let { CreateUserValidator, validationResult } = require('../utils/validatorHandler')
let userModel = require("../schemas/users");
let userController = require('../controllers/users')
let { CheckLogin, CheckRole } = require('../utils/authHandler')
let roleModel = require('../schemas/roles')
let cartModel = require('../schemas/carts')
let { sendUserPasswordMail, verifyMailConfig } = require('../utils/mailHandler')
let crypto = require('crypto')
let fs = require('fs')
let xlsx = require('xlsx')
const IMPORT_USERS_FILE_PATH = "C:\\Users\\mihvu\\Downloads\\user.xlsx";

function randomPassword() {
  return crypto.randomBytes(24).toString('base64url').slice(0, 16);
}
function sleep(ms) {
  return new Promise(function (resolve) {
    setTimeout(resolve, ms);
  });
}

function readRowsFromExcel(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error("Excel file not found");
  }
  let workbook = xlsx.readFile(filePath);
  let sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error("Excel file has no sheet");
  }
  let rows = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  return rows.map(function (row) {
    return {
      username: String(row.username || "").trim(),
      email: String(row.email || "").trim().toLowerCase()
    }
  });
}


router.get("/", CheckLogin, CheckRole("ADMIN", "MODERATOR"), async function (req, res, next) {
  let users = await userModel
    .find({ isDeleted: false })
    .populate({
      path: 'role',
      select: 'name'
    })
  res.send(users);
});

router.get("/:id",CheckLogin,CheckRole("ADMIN"), async function (req, res, next) {
  try {
    let result = await userModel
      .find({ _id: req.params.id, isDeleted: false })
    if (result.length > 0) {
      res.send(result);
    }
    else {
      res.status(404).send({ message: "id not found" });
    }
  } catch (error) {
    res.status(404).send({ message: "id not found" });
  }
});

router.post("/", CreateUserValidator, validationResult, async function (req, res, next) {
  try {
    let newItem = await userController.CreateAnUser(
      req.body.username, req.body.password, req.body.email, req.body.role
    )
    res.send(newItem);
  } catch (err) {
    res.status(400).send({ message: err.message });
  }
});

router.post("/import-excel", async function (req, res, next) {
  try {
    await verifyMailConfig();
    let rows = readRowsFromExcel(IMPORT_USERS_FILE_PATH).filter(function (row) {
      return row.username && row.email;
    });
    if (rows.length === 0) {
      return res.status(400).send({ message: "Excel has no valid rows" });
    }

    let userRole = await roleModel.findOne({ name: "USER", isDeleted: false });
    if (!userRole) {
      return res.status(400).send({ message: "Role USER not found" });
    }

    let created = [];
    let skipped = [];
    for (let row of rows) {
      let existed = await userModel.findOne({
        isDeleted: false,
        $or: [{ username: row.username }, { email: row.email }]
      });
      if (existed) {
        skipped.push({ username: row.username, email: row.email, reason: "already exists" });
        continue;
      }

      let plainPassword = randomPassword();
      let user = await userController.CreateAnUser(
        row.username,
        plainPassword,
        row.email,
        userRole._id
      );
      await cartModel.create({ user: user._id });
      await sendUserPasswordMail(row.email, row.username, plainPassword);
      await sleep(1200);
      created.push({ username: row.username, email: row.email });
    }

    res.send({
      total: rows.length,
      created: created.length,
      skipped: skipped.length,
      skippedRows: skipped
    });
  } catch (err) {
    res.status(400).send({ message: err.message });
  }
});

router.put("/:id", async function (req, res, next) {
  try {
    let id = req.params.id;
    let updatedItem = await
      userModel.findByIdAndUpdate(id, req.body, { new: true });

    if (!updatedItem) return res.status(404).send({ message: "id not found" });

    let populated = await userModel
      .findById(updatedItem._id)
    res.send(populated);
  } catch (err) {
    res.status(400).send({ message: err.message });
  }
});

router.delete("/:id", async function (req, res, next) {
  try {
    let id = req.params.id;
    let updatedItem = await userModel.findByIdAndUpdate(
      id,
      { isDeleted: true },
      { new: true }
    );
    if (!updatedItem) {
      return res.status(404).send({ message: "id not found" });
    }
    res.send(updatedItem);
  } catch (err) {
    res.status(400).send({ message: err.message });
  }
});

module.exports = router;