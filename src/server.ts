// Import the 'express' module
import express from "express";
import { sharepointWriter } from "./sharepoint/sharepointHelper";
import { generateExcel } from "./sharepoint/ExcelGenerator";

// Create an Express application
const app = express();

// Set the port number for the server
const port = 8080;

// Define a route for the root path ('/')
app.get("/", async (req, res) => {
  // Encode file name because of special characters (here "#") because the fileName goes in the API URL. You will hase a content-type error if the file name can't be read by the API.
  // Make sure you have the right file extension
  const fileName = `${encodeURIComponent("#ID")}.xlsx`;

  const sharepoint = await sharepointWriter(await generateExcel(), fileName);

  console.debug(sharepoint);
  return res.json(sharepoint);
  // Send a response to the client
  // res.send("Hello, TypeScript + Node.js + Express!");
});

// Start the server and listen on the specified port
app.listen(port, () => {
  // Log a message when the server is successfully running
  console.log(`Server is running on http://localhost:${port}`);
});
