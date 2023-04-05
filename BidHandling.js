"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.BidHandling = void 0;
var fs = require("fs");
var AdmZip = require('adm-zip');
var buffer = require("stream/consumers").buffer;
var convert = require("xml-js");
var filename = undefined;
var BidHandling = /** @class */ (function () {
    function BidHandling() {
    }
    BidHandling.BidToJson = function () {
        var copiedFolder = "C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid";
        var bidFile = fs.readdirSync(copiedFolder);
        var myFile = bidFile.filter(function (file) { return file.substring(file.length - 4, file.length).toLowerCase() === ".bid"; })[0];
        filename = myFile.substring(0, myFile.length - 4);
        fs.copyFileSync(copiedFolder + "\\" + myFile, copiedFolder + "\\" + filename + ".zip");
        fs.rmSync(copiedFolder + "\\" + myFile);
        var zip = new AdmZip(copiedFolder + "\\" + filename + ".zip");
        zip.extractAllTo(copiedFolder, true);
        bidFile = fs.readdirSync(copiedFolder);
        myFile = bidFile.filter(function (file) { return file.substring(file.length - 4, file.length).toLowerCase() === ".bid"; })[0];
        var text = fs.readFileSync(copiedFolder + "\\" + myFile, 'utf-8');
        var decodeValue = Buffer.from(text, 'base64');
        text = decodeValue.toString('utf-8');
        fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.xml", text);
        var xml = fs.readFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.xml", 'utf-8');
        var json = convert.xml2json(xml, { compact: true, spaces: 4 });
        fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.json", json);
        fs.rmSync(copiedFolder + "\\" + filename + ".zip");
        fs.rmSync(copiedFolder + "\\" + myFile);
        fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.xml");
    };
    BidHandling.JsonToBid = function () {
        var resultFilePath = "C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\OutputDataFromBID.json";
        var json = fs.readFileSync(resultFilePath, 'utf-8');
        var xml = convert.json2xml(json, { compact: true, ignoreComment: true, space: 4 });
        fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\Result_Xml.xml", xml);
        var text = fs.readFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\Result_Xml.xml", 'utf-8');
        var encodeValue = Buffer.from(text, 'utf-8');
        text = encodeValue.toString('base64');
        fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\XmlToBID.BID", text);
        var zip = new AdmZip();
        zip.addLocalFile("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\XmlToBID.BID");
        zip.writeZip("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".zip");
        fs.copyFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".zip", "C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".BID");
        fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".zip");
        fs.rmSync(resultFilePath);
        fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\Result_Xml.xml");
        fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\XmlToBID.BID");
    };
    return BidHandling;
}());
exports.BidHandling = BidHandling;
// export function BidToJson(){
//     const copiedFolder : string = "C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid";
//     let bidFile = fs.readdirSync(copiedFolder);
//     let myFile = bidFile.filter(file => file.substring(file.length - 4, file.length).toLowerCase() === ".bid")[0];
//     filename = myFile.substring(0, myFile.length - 4);
//     fs.copyFileSync(copiedFolder + "\\" + myFile, copiedFolder + "\\" + filename + ".zip");
//     fs.rmSync(copiedFolder + "\\" + myFile);
//     let zip = new AdmZip(copiedFolder + "\\" + filename + ".zip");
//     zip.extractAllTo(copiedFolder, true);
//     bidFile = fs.readdirSync(copiedFolder);
//     myFile = bidFile.filter(file => file.substring(file.length - 4, file.length).toLowerCase() === ".bid")[0];
//     let text = fs.readFileSync(copiedFolder + "\\" + myFile, 'utf-8');
//     const decodeValue = Buffer.from(text, 'base64');
//     text = decodeValue.toString('utf-8');
//     fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.xml", text);
//     const xml = fs.readFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.xml", 'utf-8');
//     const json = convert.xml2json(xml, {compact : true, spaces : 4});
//     fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.json", json);
//     fs.rmSync(copiedFolder + "\\" + filename + ".zip");
//     fs.rmSync(copiedFolder + "\\" + myFile);
//     fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.xml");
// }
// export function JsonToBid(){
//     const resultFilePath = "C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\OutputDataFromBID.json";
//     const json = fs.readFileSync(resultFilePath, 'utf-8');
//     const xml = convert.json2xml(json, {compact : true, ignoreComment : true, space : 4});
//     fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\Result_Xml.xml", xml);
//     let text = fs.readFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\Result_Xml.xml", 'utf-8');
//     const encodeValue = Buffer.from(text, 'utf-8');
//     text = encodeValue.toString('base64');
//     fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\XmlToBID.BID", text);
//     let zip = new AdmZip();
//     zip.addLocalFile("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\XmlToBID.BID");
//     zip.writeZip("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".zip");
//     fs.copyFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".zip", "C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".BID");
//     fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".zip");
//     fs.rmSync(resultFilePath);
//     fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\Result_Xml.xml");
//     fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\XmlToBID.BID");
// }
// BidHandling.BidToJson();
//JsonToBid();
