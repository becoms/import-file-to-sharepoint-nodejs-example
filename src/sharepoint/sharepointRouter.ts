import { generateExcel } from "./ExcelGenerator";
import { sharepointWriter } from "./sharepointHelper";
import express from "express";

const SharepointRouter = express.Router();

SharepointRouter.route("/").get(async (req, res) => {
  // Encode file name because of special characters (here "#") because the fileName goes in the API URL. You will hase a content-type error if the file name can't be read by the API.
  // Make sure you have the right file extension
  const fileName = `${encodeURIComponent("#test")}.xlsx`;

  // If you have a relative path with your file, uncomment next line
  // let buffer = fs.readFileSync("yourFilePath");

  //  If you have a relative path with your file, replace "generateExcel()" by "buffer" in next line.
  const sharepoint = await sharepointWriter(await generateExcel(), fileName);

  console.debug(sharepoint);
  return res.json(sharepoint);
});

export default SharepointRouter;
