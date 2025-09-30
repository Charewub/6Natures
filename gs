/**
 * Serves the HTML file for the report.
 * This script reads data from the active Google Sheet and generates a report for all entries.
 *
 * @return {HtmlOutput} The report HTML.
 */
function doGet() {
  const SPREADSHEET_ID = "1t6TMCtLF-X5lFESFCZkA40kFoVncb4x6W6JQ5R8usWI";
  const SHEET_NAME = "Result";
  
  const groupStyle = {
    "พลังแห่งวิสัยทัศน์ (Visionary Power)": {
      backgroundColor: "#6f42c1",
      textColor: "white",
      iconUrl: "https://s.imgz.io/2025/09/18/Visione9072d69abfbe932.png",
      definition: "คือพลังที่ตอบคำถามที่สำคัญที่สุดว่า \"เรากำลังจะไปที่ไหน และทำไมที่นั่นจึงสำคัญ\" พลังนี้ไม่ได้มองแค่ปัจจุบัน แต่คือการวาดภาพอนาคตที่น่าตื่นเต้นและชัดเจนจนจับต้องได้ พร้อมทั้งวางแผนที่เชิงกลยุทธ์เสมือนเป็น 'เข็มทิศ' ที่เที่ยงตรง เพื่อให้แน่ใจว่าทุกย่างก้าวของการทำงานมีความหมายและมุ่งไปสู่จุดหมายเดียวกัน"
    },
    "พลังแห่งการสร้างสรรค์ (Creative Power)": {
      backgroundColor: "#ffc107",
      textColor: "#212529",
      iconUrl: "https://s.imgz.io/2025/09/18/Creative6130d3fda7d60935.png",
      definition: "คือพลังที่กล้าตั้งคำถามกับสิ่งที่เป็นอยู่ว่า \"แล้วถ้ามีวิธีที่ดีกว่านี้ล่ะ?\" พลังนี้คือการท้าทายความจำเจและคิดนอกกรอบ เปรียบเสมือน 'หลอดไฟ' ที่ส่องสว่างขึ้นมาในความมืด มันคือการจุดประกายไอเดียใหม่ๆ, ค้นพบทางออกที่ไม่เคยมีใครเห็น, และผลักดันให้เกิดนวัตกรรมที่สร้างความแตกต่างและนำพาองค์กรไปสู่การเติบโตที่ก้าวกระโดด"
    },
    "พลังแห่งการขับเคลื่อน (Driving Power)": {
      backgroundColor: "#dc3545",
      textColor: "white",
      iconUrl: "https://s.imgz.io/2025/09/18/Drivingd1af403aa952dea6.png",
      definition: "คือพลังที่เปลี่ยนวิสัยทัศน์ให้กลายเป็นความจริง มันคือแรงขับเคลื่อนที่ไม่เคยหยุดนิ่งซึ่งถามว่า \"เราจะทำให้มันเกิดขึ้นได้อย่างไร และเมื่อไหร่?\" พลังนี้เปรียบเสมือน 'เครื่องยนต์' ที่ทรงพลังซึ่งแปลงแผนการที่สวยหรูให้กลายเป็นผลลัพธ์ที่จับต้องได้ มันคือการลงมือทำ, การจัดการอย่างมีประสิทธิภาพ, และความมุ่งมั่นที่จะผลักดันงานให้สำเร็จลุล่วง"
    },
    "พลังแห่งการเชื่อมโยง (Connecting Power)": {
      backgroundColor: "#007bff",
      textColor: "white",
      iconUrl: "https://s.imgz.io/2025/09/18/Connect4d793ab1afbb942b.png",
      definition: "คือพลังที่สร้าง \"ความร่วมมือ\" ให้เกิดขึ้นอย่างแท้จริง มันตอบคำถามว่า \"เราจะทำงานร่วมกันให้ดีที่สุดได้อย่างไร?\" พลังนี้เปรียบเสมือน 'เครือข่ายอินเทอร์เน็ต' ที่เชื่อมโยงทุกคนเข้าไว้ด้วยกัน ทำให้เกิดการสื่อสารที่ไร้รอยต่อ, ความเข้าอกเข้าใจ, และความไว้วางใจ ซึ่งเป็นรากฐานสำคัญในการสร้างทีมที่แข็งแกร่งและวัฒนธรรมองค์กรที่ยอดเยี่ยม"
    },
    "พลังแห่งคุณธรรม (Integrity Power)": {
      backgroundColor: "#737171e7",
      textColor: "white",
      iconUrl: "https://s.imgz.io/2025/09/18/Integrityc10fc536a70cc345.png",
      definition: "คือพลังที่เป็นแกนหลักทางจริยธรรมซึ่งคอยถามว่า \"สิ่งที่เรากำลังทำนั้นถูกต้องหรือไม่?\" พลังนี้เปรียบเสมือน 'ตราชั่ง' ที่เที่ยงตรง คอยสร้างความสมดุลระหว่างผลประโยชน์กับหลักการ มันคือความซื่อสัตย์, ความรับผิดชอบ, และความเป็นธรรมที่สร้างความไว้วางใจจากทั้งภายในและภายนอกองค์กร ทำให้ความสำเร็จนั้นยั่งยืนและน่าภาคภูมิใจ"
    },
    "พลังแห่งการยืนหยัด (Resilient Power)": {
      backgroundColor: "#28a745",
      textColor: "white",
      iconUrl: "https://s.imgz.io/2025/09/18/Resilientfed0a723f055609c.png",
      definition: "คือพลังที่แสดงศักยภาพสูงสุดในยามวิกฤต มันคือคำตอบของคำถามที่ว่า \"เมื่อเราล้มลง เราจะลุกขึ้นสู้อย่างไร?\" พลังนี้เปรียบเสมือน 'ต้นหลิว' ที่สามารถโอนอ่อนผ่อนตามพายุแต่ไม่เคยหักโค่น มันคือความสามารถในการปรับตัว, การฟื้นตัวจากความล้มเหลว, และการรักษาความสงบภายใต้แรงกดดัน เพื่อให้องค์กรสามารถผ่านพ้นทุกอุปสรรคและกลับมาแข็งแกร่งกว่าเดิม"
    }
  };

  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) {
    return HtmlService.createHtmlOutput('Sheet not found: ' + SHEET_NAME);
  }

  const allData = sheet.getDataRange().getValues();
  if (allData.length <= 1) {
    return HtmlService.createHtmlOutput("<h1>ไม่พบข้อมูลสำหรับสร้างรายงาน</h1>");
  }

  const headerRow = allData[0].map(header => header.toString().trim());
  const dataRows = allData.slice(1);
  
  const columnIndices = {
    employeeId: headerRow.indexOf("EmployeeID"),
    employeeName: headerRow.indexOf("EmployeeName"),
    dominantGroup: headerRow.indexOf("Dominant_Power_Group"),
    top1: headerRow.indexOf("Top_Strength_1"),
    top2: headerRow.indexOf("Top_Strength_2"),
    top3: headerRow.indexOf("Top_Strength_3"),
    top4: headerRow.indexOf("Top_Strength_4"),
    top5: headerRow.indexOf("Top_Strength_5")
  };

  for (const key in columnIndices) {
    if (columnIndices[key] === -1) {
      throw new Error(`ไม่พบชื่อคอลัมน์ "${key}" โปรดตรวจสอบว่าชื่อคอลัมน์ใน Sheet ตรงกับโค้ด`);
    }
  }

  const reportsData = dataRows.map(row => {
    const dominantGroup = row[columnIndices.dominantGroup];
    const style = groupStyle[dominantGroup] || {
      backgroundColor: "#ececdf",
      textColor: "black",
      iconUrl: "",
      definition: "No definition found."
    };
    
    const top5 = [
      row[columnIndices.top1],
      row[columnIndices.top2],
      row[columnIndices.top3],
      row[columnIndices.top4],
      row[columnIndices.top5],
    ].filter(Boolean);
    
    return {
      employeeId: row[columnIndices.employeeId],
      employeeName: row[columnIndices.employeeName],
      dominantGroup,
      top5,
      style
    };
  });

  const template = HtmlService.createTemplateFromFile('printable-output');
  template.reportsData = reportsData;

  return template.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle("Signature Strengths Report");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Report Exporter')
    .addItem('Generate All Reports (Printable)', 'doGet')
    .addToUi();
}
