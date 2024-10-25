import readXlsxFile from "read-excel-file/node"
import express from 'express'
import { createRequire } from "module";
import fs, { readFileSync } from "fs";
const require = createRequire(import.meta.url);
const Dymo = require('dymojs');
import inquirer from "inquirer";
import { start } from "repl";

const questions = [
    {
        type: 'input',
        name: 'jobNum',
        message: 'Enter job number: ',
    },
];


const dymo = new Dymo();
var JobNumber;
var Part;
var Legacy;
var CustPO;
var date;
var Description;
var CustName; 
var OpenQty;
var data;
var labelXml;
var counter = 2;


const app = express();

const schema = {
    "Job Number": {
        prop:"JobNumber",
        data: Number
    },
    "Part": {
        prop:"Part",
        data:String
    },
    "Customer PO":{
        prop: "CustomerPO",
        data: String
    },
    "Description":{
        prop: "Description",
        data: String
    },
    "Customer Name":{
        prop: "CustomerName",
        data: String
    },
    "Ship By":{
        prop: "ShipBy",
        data: String
    },
    "Open Qty":{
        prop: "OpenQty",
        data: String
    }
}

const searchSchema = {
    "Part":{
        prop:"Part",
        data:String
    },
    "Memo":{
        prop:"Legacy",
        data: String
    }
}

inquirer.prompt(questions).then(answers => {
    JobNumber = answers.jobNum;
    console.log(answers.jobNum);
    startRead();
})

function startRead() {
    if (JobNumber !== undefined){
        console.log("Passed check 1");
        readXlsxFile("./jobData/Orders to Ship.xlsx", {schema:schema}).then(async rows => {
            for (let i = 0; i < 500; i++) {
                try{
                    if (rows.rows[i].JobNumber === JobNumber) {
                        console.log("Passed check 2");
                        JobNumber = rows.rows[i].JobNumber;
                        Part = rows.rows[i].Part;
                        CustPO = rows.rows[i].CustomerPO;
                        Description = rows.rows[i].Description;
                        CustName = rows.rows[i].CustomerName;
                        OpenQty = rows.rows[i].OpenQty;
                        var nowDate = new Date(rows.rows[i].ShipBy);
                        date = (nowDate.getMonth()+1)+'/'+nowDate.getUTCDate()+'/'+nowDate.getFullYear(); 
                        break;
                    }
                }
                catch(err){
                    console.log("Error: " + err);
                }
              }
        })
        readXlsxFile("./jobData/Search Results.xlsx", {schema:searchSchema}).then(async rows => {
            for (let i = 0; i < 500; i++) {
                if (rows.rows[i].Part === Part.toString()) {
                    Legacy = rows.rows[i].Legacy;
                    break;
                }
              }
              print();
        })
    }
}

async function print(){
    //console.log(`Job#:${JobNumber} Epicor:${Part} Legacy:${Legacy} PO:${CustPO} Date:${date} Customer:${CustName} Desc:${Description} Qty:${OpenQty}`);
    labelXml = `<?xml version="1.0" encoding="utf-8"?>
<DieCutLabel Version="8.0" Units="twips" MediaType="Default">
	<PaperOrientation>Landscape</PaperOrientation>
	<Id>Shipping4x6</Id>
	<IsOutlined>false</IsOutlined>
	<PaperName>1744907 4 in x 6 in</PaperName>
	<DrawCommands>
		<RoundRectangle X="0" Y="0" Width="5918.4" Height="9038.4" Rx="270" Ry="270" />
	</DrawCommands>
	<ObjectInfo>
		<TextObject>
			<Name>TEXT</Name>
			<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />
			<BackColor Alpha="0" Red="255" Green="255" Blue="255" />
			<LinkedObjectName />
			<Rotation>Rotation0</Rotation>
			<IsMirrored>False</IsMirrored>
			<IsVariable>False</IsVariable>
			<GroupID>-1</GroupID>
			<IsOutlined>False</IsOutlined>
			<HorizontalAlignment>Center</HorizontalAlignment>
			<VerticalAlignment>Middle</VerticalAlignment>
			<TextFitMode>AlwaysFit</TextFitMode>
			<UseFullFontHeight>True</UseFullFontHeight>
			<Verticalized>False</Verticalized>
			<StyledText>
				<Element>
					<String xml:space="preserve">MCL
</String>
					<Attributes>
						<Font Family="Arial" Size="12" Bold="True" Italic="False" Underline="False" Strikeout="False" />
						<ForeColor Alpha="255" Red="0" Green="0" Blue="0" HueScale="100" />
					</Attributes>
				</Element>
				<Element>
					<String xml:space="preserve">${Description}
</String>
					<Attributes>
						<Font Family="Arial" Size="12" Bold="False" Italic="False" Underline="False" Strikeout="False" />
						<ForeColor Alpha="255" Red="0" Green="0" Blue="0" HueScale="100" />
					</Attributes>
				</Element>
				<Element>
					<String xml:space="preserve">PART# ${Legacy}
EPICOR# ${Part}
PO# ${CustPO}
JOB # ${JobNumber}
 QTY: ${OpenQty}
${date}</String>
					<Attributes>
						<Font Family="Arial" Size="12" Bold="True" Italic="False" Underline="False" Strikeout="False" />
						<ForeColor Alpha="255" Red="0" Green="0" Blue="0" HueScale="100" />
					</Attributes>
				</Element>
			</StyledText>
		</TextObject>
		<Bounds X="353.923076923077" Y="76.0000000000017" Width="8418.46153846154" Height="4320" />
	</ObjectInfo>
	<ObjectInfo>
		<ImageObject>
			<Name>GRAPHIC_1</Name>
			<ForeColor Alpha="255" Red="0" Green="0" Blue="0" />
			<BackColor Alpha="0" Red="255" Green="255" Blue="255" />
			<LinkedObjectName />
			<Rotation>Rotation0</Rotation>
			<IsMirrored>False</IsMirrored>
			<IsVariable>False</IsVariable>
			<GroupID>-1</GroupID>
			<IsOutlined>False</IsOutlined>
			<Image>iVBORw0KGgoAAAANSUhEUgAAAPEAAABuCAMAAADBJQWtAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAGtUExURQAAABBEeRJHfQ9FfA9FexFFfBFFexFGexBGexBGfAs5fw9GfAxCeQtCeRJEexFHfRJFexBFegAzfw9FeRFFfQA/PxJDfbKysqeqrKarrAA4cRBHfZ+fn6WnqaapqQAAVQ5GeKWoqg9DehFGeqerqxBFfA9Eewo/f6WqqqaoqKioqAAzZqeoqgxIeaSorA1DeAA/ahFGfaaoqqaqrKSnq6Soq6eoqxFEfaWpqw9CfRFEfKarq6WprBBGeaapqgAAfxBGepmZmaapqwxCfRBFeww/fxBHfBBEehFGfKSpq5SUqqaqrgA6dQ9EfBFEdw5CegA2bQ5Be6WoqxFHewAccaSprAA/fww/eaenqqepq6enpw5De5GRtqeprJSqqqaoq6WorKamqaWnqqWpqX9/f6qqqqGhqhFEeg1FfQBEdwk/dgAvbwA/cg9FehJCeRBEew9Ceg44cQ9Geg5DfAtHfA1GexBDexFHfAw/cg9EfRFEexBEfA9Dew1Degs3eQtGfQAqVRBDfAdBdAAzdw8/bw5FewhEdwA8eQw8eQAucw9GexBGfQ9FfQ5DfgAAAPwSaTQAAACPdFJOU/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8A8W0UKQAAAAlwSFlzAAAh1QAAIdUBBJy0nQAADjZJREFUeF7tXUty47gSfKELeOc+hU/SETpC72Yxm7mJfOaXmZUFVJGgTbXlmQmPExAI1I8oJEHJbZn9v9d/A67XW6sWfwr+8YyV4g6UflLe24zLQrdq9QnsXKNa2wC5E3x6eSr1cnuSdEf2DFerlSexzXix3sLpsEcBrJ7gZG83pPfUgOHlcqGUXrCxubAObuVJ7DiO87YlRz0dlklsfFW3KxbsIrvnOKGBZAMavUTWdgFidu+Efgc7jhVzj+vJlG83zGGB7g5+wS75zeSASDURMtiQaLutZ/cY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</Image>
			<ScaleMode>Uniform</ScaleMode>
			<BorderWidth>0</BorderWidth>
			<BorderColor Alpha="255" Red="0" Green="0" Blue="0" />
			<HorizontalAlignment>Center</HorizontalAlignment>
			<VerticalAlignment>Bottom</VerticalAlignment>
		</ImageObject>
		<Bounds X="2925" Y="4291.15384615385" Width="3433.84615384615" Height="1513.84615384615" />
	</ObjectInfo>
</DieCutLabel>
`;
dymo.renderLabel(labelXml).then(imageData => {
    fs.writeFile("./labels/images/label" + counter + ".png", imageData, "base64", (err) => {
        if (err) console.log(err);
        else {
          console.log("File written successfully\n");
          counter++;
        }
  });
});
    data = {
        labelXml: labelXml
    }
    dymo.getStatus();
    // dymo.print("DYMO LabelWriter 4XL", labelXml);
    startServer();
}
async function startServer(){
    app.use(express.static('public'));
    app.use(express.json());
    app.use(express.urlencoded({ extended: true}));


    app.get('/', (req, res)=>{
        res.sendFile(path.join(__dirname, "/public/index.html"));
    })
    app.get('/info' , (req,res)=>{
       // This will send the JSON data to the client.
        res.status(200).json(data)
    })
    
    app.post("/api/users", (req, res) => {
    //   fs.writeFile("out.png", req.body.image, "base64", (err) => {
    //     if (err) console.log(err);
    //     else {
    //       console.log("File written successfully\n");
    //     }
    //   });
        
    res.send(req.body);
});
    
    // Server setup
    app.listen(5500 , ()=>{
        console.log("Server started")
    });
}

