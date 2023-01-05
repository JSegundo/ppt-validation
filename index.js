var xmldoc = require("xmldoc")
const AdmZip = require("adm-zip")
const xml2js = require("xml2js")

class Slide {
  constructor(slideData, styles) {
    this.slideData = slideData
    this.styles = styles
  }
}

function slidesXmlWithADMZIP() {
  // Open the PowerPoint file
  const zip = new AdmZip("./presentationtest.pptx")
  // Extract the contents of the 'ppt/slides' folder
  // no me funciona:
  // const slidesXml = zip.readAsText("ppt/slides/slide5.xml")
  // console.log(slidesXml)
  // Print the XML data to the console
  var zipEntries = zip.getEntries() // an array of ZipEntry records

  zipEntries.forEach(function (zipEntry) {
    //console.log(zipEntry.toString()) // outputs zip entries information
    if (
      zipEntry.entryName.includes("ppt/slides") &&
      zipEntry.entryName.endsWith(".xml")
    ) {
      //     // transforma en string
      let xmldata = zipEntry.getData().toString("utf8")
      //     //limpia el zipentry, saca espacios vacios
      let cleanXmlData = xmldata.replace("\ufeff", "")
      //  parseXMLData(cleanXmlData)
      parseXmlData2(cleanXmlData)
    }
  })
}

function parseXmlData2(xmldata) {
  xml2js.parseString(xmldata, (err, xmlDoc) => {
    if (err) {
      console.error(err)
      return
    }

    // Access the data in the XML document

    const sldElement = xmlDoc["p:sld"]
    const spTreeElement = sldElement["p:cSld"][0]["p:spTree"][0]
    const spElements = spTreeElement["p:sp"]

    // Array to store the slides
    const slides = []

    // Iterate over the <p:sp> elements
    spElements.forEach((spElement) => {
      // Array to store the styles for this slide
      const styles = []

      // Access the text in the <p:sp> element
      const txBodyElement = spElement["p:txBody"]
      if (txBodyElement && txBodyElement[0] && txBodyElement[0]["a:p"]) {
        const pElements = txBodyElement[0]["a:p"]

        // Iterate over the <a:p> elements and access the font type and size
        pElements.forEach((pElement) => {
          // Access the font type and size
          const endParaRPrElement = pElement["a:endParaRPr"][0]
          const latinElement = endParaRPrElement["a:latin"]
          const typefaceAttr = latinElement[0]["$"]["typeface"]
          const szAttr = endParaRPrElement["$"]["sz"]

          //  Store the font type (font family) and size in the styles array
          styles.push({
            fontFamily: typefaceAttr,
            fontSize: szAttr,
          })
        })
      }
      const slide = new Slide(spElements, styles)
      // Create a new slide object with the slide data and styles

      // Add the slide object to the array of slides
      slides.push(slide)
      // console.log(slides)
    })
    return slides
  })
}

console.log(slidesXmlWithADMZIP())
// parseXMLData()
// reading archives

//
// xxxxxxxxxxxxxxxxx   dimensiones y posición de formas
// Access the style information in the <p:sp> element
// const spPrElement = spElement["p:spPr"]
// const xfrmElement = spPrElement[0]["a:xfrm"]
// const offAttr = xfrmElement[0]["a:off"]
// const extAttr = xfrmElement[0]["a:ext"] xxxxxxxxxxxxxx
//
// Print the style (dimensions and position of shapes) information to the console
// console.log(offAttr, extAttr)
