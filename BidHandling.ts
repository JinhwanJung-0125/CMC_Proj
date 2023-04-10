const fs = require("fs");
const AdmZip = require('adm-zip');
const { buffer } = require("stream/consumers");
const convert = require("xml-js");
let filename = undefined;

export class BidHandling {
    public static BidToJson() {
        const copiedFolder: string = "C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid"; // EmptyBid폴더 주소 저장 / 폴더 경로 수정 (23.02.02)
        let bidFile = fs.readdirSync(copiedFolder);
        let myFile = bidFile.filter(file => file.substring(file.length - 4, file.length).toLowerCase() === ".bid")[0]; // 확장자가 .bid인 파일을 찾기
        filename = myFile.substring(0, myFile.length - 4); // 확장자를 뺀 파일의 이름

        fs.copyFileSync(copiedFolder + "\\" + myFile, copiedFolder + "\\" + filename + ".zip"); // 확장자를 .bid에서 .zip으로 교체
        fs.rmSync(copiedFolder + "\\" + myFile); // 기존의 .bid 파일은 삭제

        let zip = new AdmZip(copiedFolder + "\\" + filename + ".zip");
        zip.extractAllTo(copiedFolder, true); // .zip파일 압축 해제

        bidFile = fs.readdirSync(copiedFolder);
        myFile = bidFile.filter(file => file.substring(file.length - 4, file.length).toLowerCase() === ".bid")[0]; // 압축 해제되어 나온 .bid파일 찾기

        let text = fs.readFileSync(copiedFolder + "\\" + myFile, 'utf-8'); // 나온 .bid 파일의 텍스트 읽기
        const decodeValue = Buffer.from(text, 'base64'); // base64코드 디코딩
        text = decodeValue.toString('utf-8'); // 디코딩 되어 나온 텍스트 저장

        fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.xml", text); // xml파일 형식으로 텍스트 쓰기

        const xml = fs.readFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.xml", 'utf-8');
        const json = convert.xml2json(xml, { compact: true, spaces: 4 }); // xml파일을 json으로 교체

        fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.json", json); // json파일 형식으로 다시 쓰기

        //=======이 과정에서 만들어진 파일들은 전부 삭제=======
        fs.rmSync(copiedFolder + "\\" + filename + ".zip");
        fs.rmSync(copiedFolder + "\\" + myFile);
        fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\OutputDataFromBID.xml");

        // Setting.GetData();
    }

    public static JsonToBid() {
        const resultFilePath = "C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\OutputDataFromBID.json";
        const json = fs.readFileSync(resultFilePath, 'utf-8'); // json파일 읽기
        const xml = convert.json2xml(json, { compact: true, ignoreComment: true, space: 4 }); // json파일을 xml파일로 교체

        fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\Result_Xml.xml", xml); // xml파일 형식으로 다시 쓰기

        let text = fs.readFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID\\Result_Xml.xml", 'utf-8'); // xml파일 내용 읽기
        const encodeValue = Buffer.from(text, 'utf-8');
        text = encodeValue.toString('base64'); // base64로 인코딩

        fs.writeFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\XmlToBID.BID", text); // 인코딩된 텍스트 XmlToBID.BID파일에 쓰기

        let zip = new AdmZip();
        zip.addLocalFile("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\XmlToBID.BID");
        zip.writeZip("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".zip"); // XmlToBID.BID파일을 .zip파일로 압축

        fs.copyFileSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".zip", "C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".BID"); // 확장자명 .zip에서 .BID로 교체

        //======이 과정에서 만들어진 파일들은 전부 삭제======
        fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\" + filename + ".zip");
        fs.rmSync(resultFilePath);
        fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\Result_Xml.xml");
        fs.rmSync("C:\\\\Users\\joung\\OneDrive\\문서\\AutoBID" + "\\EmptyBid\\XmlToBID.BID");
    }
}