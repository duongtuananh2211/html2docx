var jsdom = require("jsdom");
const $ = require("jquery")(new jsdom.JSDOM().window);
import { Document, HeadingLevel, ImageRun, Packer, Paragraph, TextRun } from "docx";
import { FileChild } from "docx/build/file/file-child";
import fs from "fs";
import axios from 'axios'


const getDataImage = async (url: string) => {

   const file = await axios.get(url, {responseType: 'arraybuffer'})
    const buffer = file.data

    console.debug('buffer', buffer)

    return buffer
}

const headingMapping: {[key: string]: HeadingLevel} = {
    'h1': HeadingLevel.HEADING_1,
    'h2': HeadingLevel.HEADING_2,
    'h3': HeadingLevel.HEADING_3,
    'h4': HeadingLevel.HEADING_4,
    'h5': HeadingLevel.HEADING_5,
    'h6': HeadingLevel.HEADING_6,
}

export default class WordGenerator {
  html: string;
  loaded: boolean = false
  fileChildren: FileChild[] = []

  constructor(html: string) {
    this.html = html;
    this.load()
  }

  load = async () => {
    this.loaded = false
    
    await this.loadElements($(this.html))
    this.generate()
  }



  loadHeading = (heading: string, element: any) => {
    const text = element.text()
    console.debug('load heading', heading, text)

    this.fileChildren.push(new Paragraph({text: text, heading: headingMapping[heading]}))
  }

  loadParagraph = (element: any) => {
    const text = element.text()
        console.debug('load paragraph', text)

    this.fileChildren.push(new Paragraph({children: [new TextRun(text)]}))
  }

  loadFigure = async (elements: any) => {
    console.debug('load figure')

   await this.loadElements(elements.children())
  }

  loadImage = async (elements: any) => {


    const src = $(elements).attr('src')
    this.fileChildren.push(new Paragraph({children: [new ImageRun({
    data: await getDataImage(src),
    transformation: {
        width: 100,
        height: 100,
    },
})]}))

    console.debug('image loaded')
  }

  loadElement = async (element: any) => {
    const tag = element.prop('tagName').toLowerCase()

    switch(tag) {
        case "h1":
        case "h2":
        case "h3": 
        case "h4": 
        case "h5": 
        case 'h6':
           return this.loadHeading(tag, element);
        case 'p':
            return this.loadParagraph(element);
        case 'figure':
            return await this.loadFigure(element);
        case 'img':
            return await this.loadImage(element)

    }
  }

  loadElements = async (elements: any) => {

    if (elements.length > 1) {
        for (let i = 0; i < elements.length; i++) {
            await this.loadElements($(elements[i]))
        }
    } else {
       await this.loadElement(elements)
    }

  }

  generate = () => {

    console.debug('generate docx')

    const doc = new Document({
      sections: [
        {
          properties: {

          },
          children: this.fileChildren,
        }
      ],
      title: 'test title'
    });

    // Used to export the file into a .docx file
    Packer.toBuffer(doc).then((buffer) => {
      fs.writeFileSync("My Document.docx", buffer);
    });
  };
}
