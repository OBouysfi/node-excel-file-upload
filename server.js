const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const upload = multer({ dest: "uploads/" });

// Serve static files
app.use(express.static(path.join(__dirname)));

// Middleware to parse form data
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// BIC Mappings based on the first 3 digits of the RIB
const bicMappings = {
  "001": "BKAMMAMR",
  "002": "ARABMAMC",
  "005": "UMABMAMC",
  "007": "BCMAMAMC",
  "011": "BMCEMAMC",
  "013": "BMCIMAMC",
  "019": "BCMAMAMC",
  "021": "CDMAMAMC",
  "022": "SGMBMAMC",
  "023": "BMCIMAMC",
  "028": "CITIMAMC",
  "101": "BCPOMAMC",
  "105": "BCPOMAMC",
  "109": "BCPOMAMC",
  "117": "BCPOMAMC",
  "127": "BCPOMAMC",
  "133": "BCPOMAMC",
  "143": "BCPOMAMC",
  "145": "BCPOMAMC",
  "148": "BCPOMAMC",
  "150": "BCPOMAMC",
  "155": "BCPOMAMC",
  "157": "BCPOMAMC",
  "159": "BCPOMAMC",
  "164": "BCPOMAMC",
  "169": "BCPOMAMC",
  "172": "BCPOMAMC",
  "175": "BCPOMAMC",
  "178": "BCPOMAMC",
  "181": "BCPOMAMC",
  "190": "BCPOMAMC",
  "205": "BKAMMAMR",
  "210": "CADGMAMR",
  "225": "CNCAMAMR",
  "230": "CIHMMAMC",
  "310": "BKAMMAMR",
  "050": "CAFGMAMC",
  "350": "ABBMMAMC",
  "833": "ABBMMAMC",
  "360": "CIHMMAMC",
  "366": "BCPOMAMCYSR",
  "863": "CNCAMAMR",
};

// Serve the sample Excel file for download
app.get("/download-sample", (req, res) => {
  res.download(path.join(__dirname, "sample_excel.xlsx"), "sample_excel.xlsx");
});

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.post("/upload", upload.single("file"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "No file uploaded." });
  }

  const { msg_id, PmtInfId, ReqdExctnDt, type } = req.body;

  // Function to normalize and replace special characters
  const normalizeString = (str) => {
    return str
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .replace(/[^a-zA-Z0-9 ]/g, '');
  };

  // Function to find SWIFT CODE based on the first 3 digits of the RIB
  const findSwiftCode = (rib) => {
    if (!rib || rib.length < 3) return null; 
    const ribWithoutSpaces = rib.replace(/\s+/g, ''); // Remove spaces from rib
    const firstThreeDigits = ribWithoutSpaces.substring(0, 3); // Get first 3 digits
    return bicMappings[firstThreeDigits] || null; // Lookup in the BIC mapping
  };

  try {
    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);

    // Extract relevant columns for XML and handle BIC replacement logic
    const dynamicData = data.map((row, index) => {
      let bic = row["bic"];
      let rib = row["rib"] ? row["rib"].replace(/\s+/g, '') : ''; // Normalize RIB
      if (!bic && rib) {
        // If BIC is empty and rib is valid, try to find the SWIFT CODE based on the RIB
        bic = findSwiftCode(rib) || "UNKNOWN"; // Default to UNKNOWN if no match
      }

      return {
        EndToEndId: `${PmtInfId}-${(++index).toString().padStart(5, "0")}`,
        InstdAmtCcy: row["currency"] || "MAD", // Default currency if not provided
        CdtrNm: `${normalizeString(row["prenom"] || '')} ${normalizeString(row["nom"] || '').trimEnd()}`,
        CdtrAcctIBAN: rib || 'UNKNOWN', // Fallback to 'UNKNOWN' if rib is undefined
        CdtrAgtBIC: bic || 'UNKNOWN', // Fallback to 'UNKNOWN' if bic is undefined
        InstdAmt: row["salaire"] || 0, // Fallback to 0 if not provided
      };
    });

    // Generate the current timestamp for CreDtTm
    const uploadTime = new Date().toISOString();

    // Calculate NbOfTxs and CtrlSum
    const nbOfTxs = dynamicData.length;
    const ctrlSum = dynamicData.reduce(
      (sum, row) => sum + parseFloat(row.InstdAmt || 0),
      0
    );

    const xml = generateXML(
      dynamicData,
      uploadTime,
      nbOfTxs,
      ctrlSum,
      msg_id,
      PmtInfId,
      ReqdExctnDt,
      type
    );

    const xmlFileName = "output.xml";
    fs.writeFileSync(path.join(__dirname, xmlFileName), xml);

    res.download(path.join(__dirname, xmlFileName), xmlFileName, (err) => {
      if (err) throw err;
      fs.unlinkSync(req.file.path);
      fs.unlinkSync(path.join(__dirname, xmlFileName));
    });
  } catch (error) {
    res.status(500).json({ error: "Error processing the file." });
  }
});

// Function to generate XML from dynamic and static data
function generateXML(
  dynamicData,
  uploadTime,
  nbOfTxs,
  ctrlSum,
  msgId,
  PmtInfId,
  ReqdExctnDt,
  type
) {
  let transactions = "";
  dynamicData.forEach((row) => {
    transactions += `
        <CdtTrfTxInf>
            <PmtId>
                <EndToEndId>${row.EndToEndId}</EndToEndId>
            </PmtId>
            <Amt>
                <InstdAmt Ccy="${row.InstdAmtCcy}">${row.InstdAmt}</InstdAmt>
            </Amt>
            <CdtrAgt>
                <FinInstnId>
                    <BIC>${row.CdtrAgtBIC}</BIC>
                </FinInstnId>
            </CdtrAgt>
            <Cdtr>
                <Nm>${row.CdtrNm}</Nm>
            </Cdtr>
            <CdtrAcct>
                <Id>
                    <Othr>
                        <Id>${row.CdtrAcctIBAN}</Id>
                    </Othr>
                </Id>
                <Ccy>MAD</Ccy>
            </CdtrAcct>
            <RmtInf>
                <Ustrd>${type}</Ustrd>
            </RmtInf>
        </CdtTrfTxInf>`;
  });

  return `<?xml version="1.0" encoding="UTF-8"?>
<Document xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03">
    <CstmrCdtTrfInitn>
        <GrpHdr>
            <MsgId>AXV ${PmtInfId}</MsgId>
            <CreDtTm>${uploadTime}</CreDtTm>
            <NbOfTxs>${nbOfTxs}</NbOfTxs>
            <CtrlSum>${ctrlSum.toFixed(2)}</CtrlSum>
            <InitgPty>
                <Nm>AXV</Nm>
            </InitgPty>
        </GrpHdr>
        <PmtInf>
            <PmtInfId>${PmtInfId}</PmtInfId>
            <PmtMtd>TRF</PmtMtd>
            <BtchBookg>true</BtchBookg>
            <NbOfTxs>${nbOfTxs}</NbOfTxs>
            <CtrlSum>${ctrlSum.toFixed(2)}</CtrlSum>
            <PmtTpInf>
                <InstrPrty>NORM</InstrPrty>
            </PmtTpInf>
            <ReqdExctnDt>${ReqdExctnDt}</ReqdExctnDt>
            <Dbtr>
                <Nm>AXV</Nm>
            </Dbtr>
            <DbtrAcct>
                <Id>
                  <Othr>
                    <Id>013780100000000018617511MAD</Id>
                  </Othr>
               </Id>
             <Ccy>MAD</Ccy>
            </DbtrAcct>
            <DbtrAgt>
                <FinInstnId>
                    <BIC>BMCIMAMC</BIC>
                </FinInstnId>
            </DbtrAgt>
            <ChrgBr>SLEV</ChrgBr>
            ${transactions}
        </PmtInf>
    </CstmrCdtTrfInitn>
</Document>`;
}

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`Server started on http://localhost:${PORT}`);
});
