
// const data =
// {
//   "pages": [
//     {
//       "pageId": "1",
//       "content": [
//         {"id": "1.1", "text": "ON-SITE GREENHOUSE", "type": "header"},
//         {"id": "1.2", "text": "WE SOURCE INGREDIENTS FROM OUR ON-SITE...", "type": "subheader"},
//         {"id": "1.3", "assetType": "image", "source": "./images/page1-1.jpg"},
//         {"id": "1.4", "text": "At Paramount Events, our green-forward initiatives...", "type": "paragraph"},
//         {"id": "1.5", "text": "fresh", "type": "watermark"},
//         {"id": "1.6", "assetType": "image", "source": "./images/page1-2.jpg"},
//         {"id": "1.7", "text": "SUSTAINABILITY", "type": "header"},
//         {"id": "1.8", "text": "WE FOCUS ON SUSTAINABILITY & GREEN-CERTIFIED...", "type": "subheader"},
//         {"id": "1.9", "text": "Our sustainability efforts include composting on-site...", "type": "paragraph"}
//       ]
//     },
//     {
//       "pageId": "2",
//       "content": [
//         {"id": "2.1", "assetType": "avatar", "source": "./images/avatar.jpg"},
//         {"id": "2.2", "text": "A FLAWLESS EXPERIENCE", "type": "header"},
//         {"id": "2.3", "text": "WE UNDERSTAND THE IMPORTANCE OF PULLING...", "type": "subheader"},
//         {"id": "2.4", "text": "No matter the guest count, venue, or reason for celebrating...", "type": "paragraph"},
//         {"id": "2.5", "text": "HI, I'M CHLOE BRYNIARSKI!", "type": "header", "name": "{Chloe Bryniarski}"},
//         {"id": "2.6", "assetType": "image", "source": "./images/page2.jpg"},
//         {"id": "2.7", "text": "I'll be your catering experience manager. Based on...", "type": "paragraph"},
//         {"id": "2.8", "text": "CALL CHLOE BRYNIARSKI OR EMAIL CBRYNIARSKI@PARAMOUNTEVENTSCHICAGO.COM", "type": "text", "name": "{Chloe Bryniarski}"},
//         {"id": "2.9", "text": "Our team", "type": "watermark"}
//       ]
//     },
//     {
//       "pageId": "3",
//       "content": [
//         {"id": "3.1", "text": "An Overview of your Event", "type": "header"},
//         {"id": "3.2", "assetType": "image", "source": "./images/page3.jpg"},
//         {"id": "3.3", "text": "Client Information", "type": "subheader"},
//         {"id": "3.4", "name": "Client Name", "email": "Client Email"},
//         {"id": "3.5", "text": "Date"},
//         {"id": "3.6", "text": "Saturday, May 5, 2029"},
//         {"id": "3.7", "text": "Venue Information", "type": "subheader"},
//         {"id": "3.8", "venue": {
//           "name": "Chicago History Museum",
//           "address": "1601 N. Clark St",
//           "city": "Illinois",
//           "state": "Chicago",
//           "zip": "60614"
//         }},
//         {"id": "3.9", "text": "Sample Timeline", "type": "subheader"},
//         {"id": "3.10", "timeline": [
//           {"time": "3:00pm", "description": "Paramount Events Team Arrives for Set Up"},
//           {"time": "6:00pm", "description": "Cocktail Reception Begins"},
//           {"time": "", "description": "Bars Open"},
//           {"time": "", "description": "Hors D'ouvre Passing Begins"},
//           {"time": "6:30", "description": "Hors D'ouvre Stations Open"},
//           {"time": "7:00pm", "description": "Hors D'ouvre Passing Concludes"},
//           {"time": "8:00pm", "description": "Event Concludes"},
//           {"time": "9:00pm", "description": "Paramount Events Completes Breakdown"}
//         ]},
//         {"id": "3.11", "text": "Guest Count", "type": "subheader"},
//         {"id": "3.12", "text": "600"},
//         {"id": "3.13", "assetType": "logo", "source": "./images/logo.jpg"}
//       ]
//     }
//   ]
// }

// import data from './data.json'
const data = require('./data.json')

page1 = data['pages'][0]['content']
page2 = data['pages'][1]['content']
page3 = data['pages'][2]['content']

const pptxgen = require('pptxgenjs')

let pres = new pptxgen()

pres.defineLayout({ name:'custom', width:11.333, height:8.5 });

// Set presentation to use new layout
pres.layout = 'custom'

let slide1 = pres.addSlide();

slide1.addText(page1[0]['text'], { x: 0.5, y: 0.5, h:0.3, fontFace: "Optima", fontSize: 18, color: "000000", bold: false, italic: false, underline: false });
slide1.addText(page1[1]['text'], { x: 1, y: 1, h:0.3, fontFace: "Avenir Regular", fontSize: 11, color: "000000", bold: false, italic: false, underline: false });
slide1.addText(page1[3]['text'], { x: 1, y: 1.3, h: 3, w:4 , linespacing: 8,fontFace: "Avenir Book", fontSize: 10,wrap:true, valign:"top", color: "000000", bold: false, italic: false, underline: false });
slide1.addText(page1[6]['text'], { x: 7.25, y: 5.5, w: 3, h:1, fontFace: "Optima", fontSize: 18, color: "000000", align: "right", valign:"bottom", bold: false, italic: false, underline: false });
slide1.addText(page1[7]['text'], { x: 6.25, y: 6.7, w: 4, h:0.2, fontFace: "Avenir Regular", fontSize: 11, color: "000000", align: "right", bold: false, italic: false, underline: false });
slide1.addText(page1[8]['text'], { x: 6.25, y: 6.9, h: 1.5, w:4 , linespacing: 5, fontFace: "Avenir Book", fontSize: 10,wrap:true, align: "right", valign:"top", color: "000000", bold: false, italic: false, underline: false });
slide1.addText(page1[4]['text'], { x: 0.7, y: "50%", fontFace: "Avenir Book", fontSize: 180, color: "000000", transparency: 95});
slide1.addImage({ path: page1[2]['source'], y:0.5, x: 7.25, w:3.5, h:4.5});
slide1.addImage({ path: page1[5]['source'], y:5, x:0, w:4.6, h:3});
slide1.addShape(pres.shapes.LINE, {
  x: 10.5,
  y: 5.5,
  w: 0,
  h: 1.8,
  line: { color: "000000", width: 1}
});
slide1.addShape(pres.shapes.LINE, {
  x: 0.75,
  y: 1.3,
  w: 0,
  h: 3,
  line: { color: "000000", width: 1}
});



let slide2 = pres.addSlide();

slide2.addText(page2[1]['text'], { x: 6.5, y: 0.7, w: 3.5, h:0.3, fontFace: "Optima", fontSize: 18, color: "000000", align: "right", bold: false, italic: false, underline: false });
slide2.addText(page2[2]['text'], { x: 6, y: 1, w: 4, h:0.4, fontFace: "Avenir Regular", fontSize: 11, color: "000000", align: "right", wrap:true, bold: false, italic: false, underline: false });
slide2.addText(page2[3]['text'], { x: 6.5, y: 1.7, h: 1, w:3.5 , fontFace: "Avenir Book", fontSize: 10.5, align: "right", wrap:true, valign:"top", color: "000000", bold: false, italic: false, underline: false });
slide2.addText(page2[4]['text'], { x: 1, y: 4.25, w: 3.5, h:0.2, fontFace: "Avenir Regular", fontSize: 11, color: "000000", align: "left", bold: false, italic: false, underline: false });
slide2.addText(page2[6]['text'], { x: 1, y: 4.5, w: 3.5, h:1, fontFace: "Avenir Book", fontSize: 10.5, color: "000000", align: "left", valign: "top", wrap:true, bold: false, italic: false, underline: false });
slide2.addText(page2[7]['text'], { x: 1, y: 5.5, h: 0.5, w:4 , fontFace: "Avenir Regular", fontSize: 11,wrap:true, align: "left", valign:"top",wrap:true, color: "000000", bold: false, italic: false, underline: false });
slide2.addImage({ path: page2[0]['source'], y:1.3, x: 1.5, w:2.5, h:2.5, rounding: true});
slide2.addImage({ path: page2[5]['source'], y:2.8, x:7, w:3.7, h:3.5});

slide2.addText(page2[8]['text'], { x: 0, y: 6.2, w:"100%", h:2, fontFace: "Avenir Book",align:"center", valign:"middle", fontSize: 180, color: "000000", transparency: 95});

let slide3 = pres.addSlide();

slide3.addText(page3[0]['text'], { x: 0.5, y: 0.7, w: 3.5, h:0.3, fontFace: "Optima", fontSize: 18, color: "000000", align: "left", bold: true, italic: false, underline: false });
slide3.addText(page3[2]['text'], { x: 0.5, y: 1.3, w: 3.5, h:0.3, fontFace: "Avenir Regular", fontSize: 9, color: "000000", align: "left", bold: true, italic: false, underline: false });
slide3.addText(page3[3]['name'], { x: 0.5, y: 1.6, w: 3.5, h:0.2, fontFace: "Avenir Book", fontSize: 9, color: "000000", align: "left", bold: false, italic: false, underline: false });
slide3.addText(page3[3]['email'], { x: 0.5, y: 1.8, w: 3.5, h:0.2, fontFace: "Avenir Book", fontSize: 9, color: "000000", align: "left", bold: false, italic: false, underline: false });
slide3.addText(page3[4]['text'], { x: 0.5, y: 2.2, w: 3.5, h:0.3, fontFace: "Avenir Regular", fontSize: 9, color: "000000", align: "left", bold: true, italic: false, underline: false });
slide3.addText(page3[5]['text'], { x: 0.5, y: 2.5, w: 3.5, h:0.2, fontFace: "Avenir Book", fontSize: 9, color: "000000", align: "left", bold: false, italic: false, underline: false });
slide3.addText(page3[6]['text'], { x: 0.5, y: 2.8, w: 3.5, h:0.3, fontFace: "Avenir Regular", fontSize: 9, color: "000000", align: "left", bold: true, italic: false, underline: false });
slide3.addText(page3[7]['venue']['name'], { x: 0.5, y: 3.1, w: 3.5, h:0.2, fontFace: "Avenir Book", fontSize: 9, color: "000000", align: "left", bold: false, italic: false, underline: false });
slide3.addText(page3[7]['venue']['address'], { x: 0.5, y: 3.3, w: 3.5, h:0.2, fontFace: "Avenir Book", fontSize: 9, color: "000000", align: "left", bold: false, italic: false, underline: false });
slide3.addText(page3[7]['venue']['city'] + ", "+ page3[7]['venue']['state'] + ", "+page3[7]['venue']['zip'], { x: 0.5, y: 3.5, w: 3.5, h:0.2, fontFace: "Avenir Book", fontSize: 9, color: "000000", align: "left", bold: false, italic: false, underline: false });
slide3.addText(page3[8]['text'], { x: 0.5, y: 3.9, w: 3.5, h:0.3, fontFace: "Avenir Regular", fontSize: 9, color: "000000", align: "left", bold: true, italic: false, underline: false });

table_y = 4.2
let rows = [];

// Row One: cells will be formatted according to any options provided to `addTable()`
table = page3[9]['timeline']

for( let x=0;x < table.length-1; x++)
  { 
    table_y += 0.2
    rows.push([{ text: table[x]["time"], options: { fontFace: "Calibri (Body)", fontSize: 8, color: "000000", align: "left", color: "000000", border: [null, {color:'DDDDDD'}, {color:'DDDDDD'}, null]} }, { text: table[x]['description'], options: { fontFace: "Calibri (Body)", fontSize: 8, color: "000000", align: "left", color: "000000", border: [null, null, {color:'DDDDDD'}, null] } }]);
  }
rows.push([{ text: table[table.length-1]["time"], options: { fontFace: "Calibri (Body)", fontSize: 8, color: "000000", align: "left", color: "000000", border: [null, {color:'DDDDDD'}, null, null]} }, { text: table[table.length-1]['description'], options: { fontFace: "Calibri (Body)", fontSize: 8, color: "000000", align: "left", color: "000000"} }]);
table_y += 0.6
slide3.addTable(rows, { x: 0.5, y: 4.2, w: 5, color: "363636", colW:[1.0, 4.0]});

slide3.addText(page3[10]['text'], { x: 0.5, y: table_y, w: 3.5, h:0.3, fontFace: "Avenir Regular", fontSize: 11, color: "000000", align: "left", bold: true, italic: false, underline: false });
slide3.addText(page3[11]['text'], { x: 0.5, y: table_y+0.3, w: 3.5, h:0.2, fontFace: "Avenir Book", fontSize: 11, color: "000000", align: "left", bold: false, italic: false, underline: false });

slide3.addImage({ path: page3[1]['source'], y:0, x: 6.533, w:4.8, h:"100%"});
slide3.addImage({ path: page3[12]['source'], y:7.3, x: 0.5, w:0.5, h:0.6});
slide3.addShape(pres.shapes.LINE, {
    x: 0.3,
    y: 0.5,
    w: 0,
    h: 7.5,
    line: { color: "4A7EBB", width: 1}
  });

pres.writeFile({ fileName: 'result.pptx' });
