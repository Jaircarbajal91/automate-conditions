interface TableDataRow {
  [key: string]: string | object[];
}

interface ConditionData {
  Condition: TableDataRow;
  Parts: TableDataRow[];
  Dispositions: TableDataRow[];
  Comments: TableDataRow[];
  Links: TableDataRow[];
  DispositionApprovals: TableDataRow | TableDataRow[];
}


function tableToObject(table: ExcelScript.Table): { [key: string]: string[] } {
  let headerRow = table.getHeaderRowRange().getTexts()[0];
  let values = table.getRangeBetweenHeaderAndTotal().getTexts();
  let obj: { [key: string]: string[] } = {};

  headerRow.forEach((header, index) => {
    obj[header] = values.map(row => row[index]);
  });

  return obj;
}

function main(workbook: ExcelScript.Workbook) {
  let tables = workbook.getTables();

  if (tables.length === 0) {
    console.log("No tables found in the workbook.");
    return;
  }

  let result: { [key: string]: ConditionData } = {};

  // Process each table
  for (let table of tables) {
    const tableName = table.getName();


    const currentSheet = workbook.getWorksheet(tableName)
    const usedRange = currentSheet.getUsedRange()
    const address = usedRange.getAddress()
    const range = address.split("!")[1]
    table.resize(range)

    const tableData: { [key: string]: string[] } = tableToObject(table);

    switch (tableName) {
      case 'Conditions':
        for (let i = 0; i < tableData['Condition Number'].length; i++) {
          const conditionNumber = tableData['Condition Number'][i];
          if (!result[conditionNumber]) {
            result[conditionNumber] = {
              Condition: {},
              Parts: [],
              Dispositions: [],
              Comments: [],
              Links: [],
              DispositionApprovals: []
            };
          }
          for (let header in tableData) {
            result[conditionNumber].Condition[header] = tableData[header][i];
          }
        }
        break;

      case 'Parts':
      case 'Dispositions':
      case 'Comments':
      case 'Links':
      case 'DispositionApprovals': {
        for (let i = 0; i < tableData['Condition Number'].length; i++) {
          const conditionNumber = tableData['Condition Number'][i];
          let row: TableDataRow = {};
          for (let header in tableData) {
            row[header] = tableData[header][i];
          }
          result[conditionNumber]?.[tableName].push(row);
        }
        break;
      }
    }
  }

  for (let key in result) {
    const dispositionApprovals = result[key].DispositionApprovals
    for (let obj of dispositionApprovals) {
      const condition: TableDataRow = result[obj["Condition Number"]]
      const disposition: TableDataRow = condition.Dispositions.find(disp => disp["Disposition Number"] == obj["Disposition Number"])
      if (disposition["Disposition Approvals"]) disposition["Disposition Approvals"].push(obj)
      else disposition["Disposition Approvals"] = [obj]
    }
    delete result[key].DispositionApprovals;
  }
  console.log(result);
  renderMainSheet(workbook, result);
}



function renderMainSheet(workbook: ExcelScript.Workbook, data: { [key: string]: ConditionData }) {
  // If "Main" sheet exists, remove it
  const existingSheet = workbook.getWorksheet("Main");
  if (existingSheet) {
    existingSheet.delete();
  }

  // Create the "Main" worksheet
  const mainSheet: ExcelScript.Worksheet = workbook.addWorksheet("Main");

  // Write Header table
  writeHeader(workbook, mainSheet)

  // write condition count in A3 cell
  const conditionCountRange = mainSheet.getRange('A5')
  const conditionCount = Object.keys(data).length
  conditionCountRange.setValue(`Conditions (${conditionCount})`)
  conditionCountRange.getFormat().getFont().setBold(true);
  conditionCountRange.getFormat().getFont().setSize(13);

  let topRow: number = 7
  // create conditions ranges
  for (const conditionNumber in data) {
    const conditionData = data[conditionNumber];
    topRow = writeCondition(workbook, mainSheet, conditionData, topRow)
    // topRow = bottomRow + 3
    // bottomRow = topRow + 100
    // topRow += 10
  }
}

function writeCondition(workbook: ExcelScript.Workbook, mainSheet: ExcelScript.Worksheet, condition: ConditionData, topRow: number) {
  const statusCell = mainSheet.getRange(`G${topRow}:H${topRow}`)
  statusCell.getFormat().getFont().setBold(true);
  statusCell.getFormat().getFont().setColor("White");
  statusCell.getFormat().getFont().setSize(13);

  const FIRSTROW = mainSheet.getRange(`${topRow}:${topRow}`)
  FIRSTROW.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  FIRSTROW.getFormat().setRowHeight(26)
  statusCell.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center)
  // console.log(condition)
  if (condition["Condition"]["Status"] === 'Open') statusCell.getFormat().getFill().setColor("#1F3864");
  if (condition["Condition"]["Status"] === 'Closed') statusCell.getFormat().getFill().setColor("#71AF84");

  statusCell.setValue(`${condition["Condition"]["Status"]}`)
  statusCell.merge();

  const rejectCellsToMerge = mainSheet.getRange(`A${topRow}:C${topRow}`)
  const rejectCellsData = mainSheet.getRange(`A${topRow}:C${topRow}`)

  rejectCellsToMerge.unmerge();
  rejectCellsData.setValue(`${condition["Condition"]["Reject Category"]} / ${condition["Condition"]["Reject Code"]}`);
  rejectCellsData.getFormat().getFont().setBold(true)
  rejectCellsData.getFormat().getFont().setSize(11)
  rejectCellsToMerge.merge();

  const conditionNumberMerge = mainSheet.getRange(`D${topRow}:E${topRow}`)
  const conditionNumberData = mainSheet.getRange(`D${topRow}:E${topRow}`)
  conditionNumberMerge.unmerge();
  conditionNumberData.setValue(`Condition ${condition["Condition"]["Condition Number"]}`);
  conditionNumberData.getFormat().getFont().setSize(13)
  conditionNumberData.getFormat().getFill().setColor('#111111')
  conditionNumberData.getFormat().getFont().setColor('White')
  conditionNumberData.getFormat().getFont().setBold(true)
  conditionNumberMerge.merge();
  conditionNumberMerge.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center)

  // SECOND ROW condition updated by
  topRow++
  const conditionUpdatedByMerge = mainSheet.getRange(`A${topRow}:H${topRow}`)
  const conditionUpdatedByData = mainSheet.getRange(`A${topRow}`)
  conditionUpdatedByData.getFormat().getFont().setColor("#6B6B6B");

  conditionUpdatedByMerge.unmerge();

  conditionUpdatedByData.setValue(`Condition updated on ${condition["Condition"]["Updated Date"].split('+')[0]}
   by ${condition["Condition"]["Updated By"]}`);

  conditionUpdatedByData.getFormat().autofitRows();

  let adjustedHeight: number = conditionUpdatedByData.getFormat().getRowHeight() - 10;


  conditionUpdatedByMerge.merge();
  conditionUpdatedByMerge.getFormat().setRowHeight(adjustedHeight);
  conditionUpdatedByMerge.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  conditionUpdatedByMerge.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);


  topRow++
  // THIRD-FOURTH ROW condition where found
  const whereFoundText = mainSheet.getRange(`A${topRow}: H${topRow}`)
  whereFoundText.setValue(`Where Found:`)
  whereFoundText.merge()
  whereFoundText.getFormat().getFont().setBold(true)
  topRow++
  const whereFoundValue = mainSheet.getRange(`A${topRow}: H${topRow}`)
  whereFoundValue.setValue(`${condition["Condition"]["Where Found"]}`)
  whereFoundValue.merge()
  whereFoundValue.getFormat().setWrapText(true)



  // FIFTH-SIXTH ROW condition Problem Source
  topRow++
  const problemSourceText = mainSheet.getRange(`A${topRow}: H${topRow}`)
  problemSourceText.setValue(`Problem Source:`)
  problemSourceText.merge()
  problemSourceText.getFormat().getFont().setBold(true)
  topRow++
  const problemSourceValue = mainSheet.getRange(`A${topRow}: H${topRow}`)
  problemSourceValue.setValue(`${condition["Condition"]["Problem Source"]}`)
  problemSourceValue.merge()
  problemSourceValue.getFormat().setWrapText(true)



  // SEVENTH-EIGHTH ROW condition CauseCode Source
  topRow++
  const causeCodeText = mainSheet.getRange(`A${topRow}: H${topRow}`)
  causeCodeText.setValue(`Cause Code:`)
  causeCodeText.merge()
  causeCodeText.getFormat().getFont().setBold(true)
  topRow++
  const causeCodeValue = mainSheet.getRange(`A${topRow}: H${topRow}`)
  causeCodeValue.setValue(`${condition["Condition"]["Problem Source"]}`)
  causeCodeValue.merge()
  causeCodeValue.getFormat().setWrapText(true)


  // NINTH ROW condition Program Source
  topRow++
  const programText = mainSheet.getRange(`A${topRow}: H${topRow}`)
  programText.setValue(`Program: ${condition["Condition"]["Custom Attribues"] || 'N/A'}`)
  programText.merge()
  programText.getFormat().getFont().setBold(true)

  // TENTH ROW condition PARTS HEADER
  topRow++
  const partsHeaderRange = mainSheet.getRange(`A${topRow}:H${topRow}`);
  partsHeaderRange.unmerge();
  partsHeaderRange.setValue(`Defective Parts (${condition["Parts"].length})`);
  partsHeaderRange.getFormat().getFont().setBold(true);
  partsHeaderRange.getFormat().getFont().setSize(13);
  partsHeaderRange.merge();
  partsHeaderRange.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  partsHeaderRange.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // 12th ROW parts
  let currentRow = topRow + 1
  for (let part of condition["Parts"]) {
    const partNumberRange = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    partNumberRange.unmerge();
    partNumberRange.getFormat().getFont().setBold(true);
    partNumberRange.setValue(`${part["Part Number"]} / ${part["Revision"]}`);
    partNumberRange.merge();

    currentRow++
    const partDescriptionRange = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    partDescriptionRange.unmerge();
    partDescriptionRange.setValue(`${part["Part Description"]}`);
    partDescriptionRange.getFormat().getFont().setColor("#6B6B6B");
    partDescriptionRange.merge();

    currentRow++
    const iTagRange = mainSheet.getRange(`A${currentRow}:D${currentRow}`)
    iTagRange.unmerge();
    iTagRange.setValue(`iTag: ${part["iTag"]}`);
    iTagRange.merge();
    const serialRange = mainSheet.getRange(`E${currentRow}:H${currentRow}`)
    serialRange.unmerge();
    serialRange.setValue(`${part["Serial / Lot Number"]}`);
    serialRange.merge();

    currentRow++
    const mustResolveByRange = mainSheet.getRange(`A${currentRow}:D${currentRow}`)
    mustResolveByRange.unmerge();
    mustResolveByRange.setValue(`Must Resolve By: ${part["Must Resolve By"]}`);
    mustResolveByRange.merge();

    currentRow++
    const locationRange = mainSheet.getRange(`A${currentRow}:D${currentRow}`)
    locationRange.unmerge();
    locationRange.setValue(`Location: ${part["Location"] || 'N/A'}`);
    locationRange.merge();

    currentRow += 2
  }
  topRow = currentRow + 1

  const isConditionHeader = mainSheet.getRange(`A${topRow}:H${topRow}`)
  isConditionHeader.unmerge();
  isConditionHeader.getFormat().getFont().setBold(true)
  isConditionHeader.setValue(`Is Condition:`);
  isConditionHeader.merge();

  topRow++
  // const isConditionRemovedHtml:string = removeHTML(condition["Condition"]["Is Condition"])
  const isCondition = mainSheet.getRange(`A${topRow}:H${topRow}`)
  isCondition.unmerge();
  isCondition.getFormat().setWrapText(true)
  const isConditionRemovedHTML: string = removeHTML(condition["Condition"]["Is Condition"])
  isCondition.setValue(`${isConditionRemovedHTML}`);
  isCondition.merge();
  isCondition.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  isCondition.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  topRow++
  const shouldBeHeader = mainSheet.getRange(`A${topRow}:H${topRow}`)
  shouldBeHeader.unmerge();
  shouldBeHeader.getFormat().getFont().setBold(true)
  shouldBeHeader.setValue(`Should Be:`);
  shouldBeHeader.merge();

  topRow++
  const shouldBeCondition = mainSheet.getRange(`A${topRow}:H${topRow}`)
  shouldBeCondition.unmerge();
  shouldBeCondition.getFormat().setWrapText(true)
  const shouldBeConditionRemovedHTML: string = removeHTML(condition["Condition"]["Should be "])
  shouldBeCondition.setValue(`${shouldBeConditionRemovedHTML}`);
  shouldBeCondition.merge();
  shouldBeCondition.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  shouldBeCondition.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);


  topRow += 2
  const dispositionHeader = mainSheet.getRange(`A${topRow}:H${topRow}`)
  dispositionHeader.unmerge();
  dispositionHeader.setValue(`Dispositions (${condition["Dispositions"].length})`);
  dispositionHeader.getFormat().getFont().setBold(true);
  dispositionHeader.getFormat().getFont().setSize(16);
  dispositionHeader.merge();
  dispositionHeader.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  dispositionHeader.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center)

  currentRow = topRow + 2
  for (let dispo of condition["Dispositions"]) {
    const dispoUpdateBy = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    dispoUpdateBy.unmerge()
    dispoUpdateBy.setValue(`Disposition Updated ${dispo["Modified Date"].split('+')[0] || dispo["Created Date"].split('+')[0]} by ${dispo["Modified By"] || dispo["Created By"]}`)
    dispoUpdateBy.getFormat().getFont().setColor("#6B6B6B");
    dispoUpdateBy.merge()
    dispoUpdateBy.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    dispoUpdateBy.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right)

    currentRow++
    const dispoNum = mainSheet.getRange(`A${currentRow}`)
    dispoNum.setValue(`Disposition ${dispo["Disposition Number"]}`);
    dispoNum.getFormat().getFont().setBold(true);

    const status = mainSheet.getRange(`F${currentRow}:H${currentRow}`)
    if (dispo["Status"] === "Completed") {
      status.getFormat().getFill().setColor("#71AF84");
    } else {
      status.getFormat().getFill().setColor("#969696");
    }
    status.unmerge()
    status.setValue(`${dispo["Status"]}`);
    status.getFormat().getFont().setBold(true);
    status.getFormat().getFont().setSize(12);
    status.getFormat().getFont().setColor('white')
    status.merge();
    status.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    status.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center)

    const dispoType = mainSheet.getRange(`C${currentRow}:E${currentRow}`)
    dispoType.unmerge()
    dispoType.setValue(`Type: ${dispo["Type"]}`);
    dispoType.merge();
    dispoType.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

    currentRow++
    const dispositionMargin = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    dispositionMargin.getFormat().setRowHeight(10)
    currentRow++
    const dispositionAssigneeHeader = mainSheet.getRange(`A${currentRow}`)
    dispositionAssigneeHeader.setValue(`Disposition Assignee:
${dispo["Disposition Assignee"] || 'N/A'}`);
    // let adjustedHeight: number = dispositionAssigneeHeader.getFormat().getRowHeight();
    // dispositionAssigneeHeader.getFormat().setRowHeight(adjustedHeight);
    dispositionAssigneeHeader.getFormat().setRowHeight(33)
    dispositionAssigneeHeader.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    dispositionAssigneeHeader.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    const requireVerification = mainSheet.getRange(`B${currentRow}:D${currentRow}`)
    requireVerification.unmerge()
    requireVerification.setValue(`Require Verification:
${dispo["Require Verification"] || 'N/A'}`);
    // let adjustedHeight: number = dispositionAssigneeHeader.getFormat().getRowHeight();
    // dispositionAssigneeHeader.getFormat().setRowHeight(adjustedHeight);
    requireVerification.merge()
    requireVerification.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    requireVerification.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    const classification = mainSheet.getRange(`E${currentRow}:G${currentRow}`)
    classification.unmerge()
    classification.setValue(`Classification:
${dispo["Classification"] || 'N/A'}`);
    // let adjustedHeight: number = dispositionAssigneeHeader.getFormat().getRowHeight();
    // dispositionAssigneeHeader.getFormat().setRowHeight(adjustedHeight);
    classification.merge()
    classification.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    classification.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    currentRow++
    const riskAssessment = mainSheet.getRange(`A${currentRow}:C${currentRow}`)
    riskAssessment.unmerge()
    riskAssessment.setValue(`Risk Assessment:
${dispo["Risk Assessment"] || 'N/A'}`);
    riskAssessment.merge()
    riskAssessment.getFormat().setRowHeight(33)
    riskAssessment.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    riskAssessment.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    if (dispo["Type"] === "Return to vendor") {
      const repair = mainSheet.getRange(`D${currentRow}:F${currentRow}`)
      repair.unmerge()
      repair.setValue(`Repair:
${dispo["Repair"]}`);
      repair.merge()
      repair.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
      repair.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
    }

    currentRow++
    const rationale = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    rationale.unmerge()
    rationale.setValue(`Rationale:
${dispo["Rationale"]}`);
    rationale.merge()
    rationale.getFormat().setWrapText(true)
    rationale.getFormat().setRowHeight(33)
    rationale.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    rationale.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    currentRow++
    if ("Disposition Approvals" in dispo) {
      const dispoApprovalHeader = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
      dispoApprovalHeader.unmerge()
      dispoApprovalHeader.setValue(`Diposition Approvals (${dispo["Disposition Approvals"].length})`)
      dispoApprovalHeader.getFormat().getFont().setBold(true);
      dispoApprovalHeader.getFormat().getFont().setSize(12);
      dispoApprovalHeader.merge()
      dispoApprovalHeader.getFormat().setRowHeight(20)
      dispoApprovalHeader.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
      dispoApprovalHeader.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      currentRow++
      const margin = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
      margin.getFormat().setRowHeight(12)
      currentRow++
      for (let approval of dispo["Disposition Approvals"]) {
        const qualification = mainSheet.getRange(`A${currentRow}`)
        qualification.setValue(`Qualification:
${approval["Qualification"]}`)
        qualification.getFormat().getFont().setBold(true);
        qualification.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);


        const approvers = mainSheet.getRange(`B${currentRow}:E${currentRow}`)
        approvers.unmerge()
        approvers.setValue(`${approval["Approvers"]}`)
        approvers.getFormat().getFont().setBold(true);
        approvers.getFormat().setWrapText(true)
        approvers.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
        approvers.merge()

        if (approval["Signature"]) {
          const signature = mainSheet.getRange(`F${currentRow}:H${currentRow}`)
          signature.unmerge()
          const splitSignature:string[] = approval["Signature"].split(',')
          const name: string = splitSignature[0]
          const time: string = splitSignature[1].split('+')[0]
          signature.setValue(`Approved: 
by ${name}
on ${time}`)
          signature.getFormat().getFont().setBold(true);
          signature.getFormat().setWrapText(true)
          signature.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
          signature.merge()
        }

        currentRow++
      }
    } else {
      const dispoApprovalHeader = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
      dispoApprovalHeader.unmerge()
      dispoApprovalHeader.setValue(`Diposition Approvals (0)`)
      dispoApprovalHeader.getFormat().getFont().setBold(true);
      dispoApprovalHeader.getFormat().getFont().setSize(12);
      dispoApprovalHeader.merge()
      dispoApprovalHeader.getFormat().setRowHeight(20)
      dispoApprovalHeader.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
      dispoApprovalHeader.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      currentRow++
    }
    currentRow++
    const instructionsHeader = mainSheet.getRange(`A${currentRow}`)
    instructionsHeader.setValue(`Disposition Instructions:`)
    instructionsHeader.getFormat().getFont().setBold(true);
    
    currentRow++
    const instructions = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    instructions.unmerge()
    const removeInstructionsHTML: string = removeHTML(dispo["Instructions"])
    instructions.setValue(`${removeInstructionsHTML}`)
    instructions.merge()
    instructions.getFormat().setWrapText(true)
    instructions.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
    instructions.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    currentRow++
    let margin = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    margin.getFormat().setRowHeight(10)

    
    currentRow++
    const executionHeader = mainSheet.getRange(`A${currentRow}`)
    executionHeader.setValue(`Execution Notes:`)
    executionHeader.getFormat().getFont().setBold(true);
    
    currentRow++
    const executionNotes = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    executionNotes.unmerge()
    const executionNotesHTML: string = removeHTML(dispo["ExecutionNotes"])
    executionNotes.setValue(`${executionNotesHTML}`)
    executionNotes.merge()
    executionNotes.getFormat().setWrapText(true)
    executionNotes.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
    executionNotes.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    currentRow++
    margin = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    margin.getFormat().setRowHeight(10)


    currentRow++
    const verificationHeader = mainSheet.getRange(`A${currentRow}`)
    verificationHeader.setValue(`Verification Notes:`)
    verificationHeader.getFormat().getFont().setBold(true);

    currentRow++
    const verificationNotes = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    verificationNotes.unmerge()
    const verificationNotesHTML: string = removeHTML(dispo["Verification Notes"])
    verificationNotes.setValue(`${verificationNotesHTML}`)
    verificationNotes.merge()
    verificationNotes.getFormat().setWrapText(true)
    verificationNotes.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
    verificationNotes.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    currentRow++
    margin = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    margin.getFormat().setRowHeight(10)

    currentRow += 1
  }
  topRow = currentRow
  const linksHeader = mainSheet.getRange(`A${topRow}:H${topRow}`)
  linksHeader.unmerge()
  linksHeader.getFormat().getFont().setSize(13)
  linksHeader.setValue(`Links (${condition["Links"].length})`)
  linksHeader.getFormat().getFont().setBold(true);
  linksHeader.merge()
  linksHeader.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  linksHeader.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  currentRow = topRow + 1
  for (let link of condition["Links"]) {
    const linkType = mainSheet.getRange(`A${currentRow}:B${currentRow}`)
    linkType.unmerge()
    linkType.setValue(`${link["Type"]}`)
    linkType.getFormat().getFont().setBold(true);
    linkType.merge()
    currentRow++

    const reference = mainSheet.getRange(`A${currentRow}:H${currentRow}`)
    reference.unmerge()
    reference.setValue(`${link["Reference"]}`)
    reference.merge()
    reference.getFormat().setWrapText(true)

    currentRow++
  }

  topRow = currentRow
  let margin = mainSheet.getRange(`A${topRow}:H${topRow}`)
  margin.getFormat().setRowHeight(10)

  topRow++
  const commentsHeader = mainSheet.getRange(`A${topRow}:H${topRow}`)
  commentsHeader.unmerge()
  commentsHeader.getFormat().getFont().setSize(13)
  commentsHeader.setValue(`Comments (${condition["Comments"].length})`)
  commentsHeader.getFormat().getFont().setBold(true);
  commentsHeader.merge()
  commentsHeader.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  commentsHeader.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  topRow++
  for (let comment of condition["Comments"]) {

    const commentInfo = mainSheet.getRange(`A${topRow}:H${topRow}`)
    commentInfo.unmerge()
    commentInfo.setValue(`Comment created on ${comment["Created On"].split('+')[0]} by ${comment["Author"]}`)
    commentInfo.getFormat().getFont().setColor('#6B6B6B')
    commentInfo.merge()
    commentInfo.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
    commentInfo.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);

    topRow++
    const commentText = mainSheet.getRange(`A${topRow}:H${topRow}`)
    commentText.unmerge()
    commentText.setValue(`${comment["Comment"]}`)
    commentText.getFormat().setWrapText(true)
    commentText.merge()
    commentText.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
    commentText.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

    topRow+=2
  }


  topRow+=2
  const endOfCondition = mainSheet.getRange(`A${topRow}:H${topRow}`)
  endOfCondition.unmerge()
  endOfCondition.setValue(`End of Condition ${condition["Condition"]["Condition Number"]}`)
  endOfCondition.getFormat().getFont().setColor("white")
  endOfCondition.merge()
  endOfCondition.getFormat().getFill().setColor('Black')
  endOfCondition.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  endOfCondition.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  topRow += 4
  return topRow
}


function removeHTML(inputString: string) {
  return inputString
    .replace(/&nbsp;/gi, ' ')      // Replace &nbsp; with space
    .replace(/&[a-z]+;/gi, ' ')   // Replace other named HTML entities with space (like &gt;, &lt;, etc.)
    .replace(/<br\s*\/?>/gi, `
    `)  // Replace <br> and its variants with line breaks
    .replace(/<[^>]*>/g, ' ')      // Remove all other HTML tags and replace them with spaces
    .replace(/\s+/g, ' ')          // Replace multiple spaces with a single space
    .replace(/\s+\n/g, `
    `)       // Remove spaces before a newline
    .replace(/\n\s+/g, `
    `)       // Remove spaces after a newline
    .trim();                       // Remove any spaces or newlines at the beginning and end
}


function writeHeader(workbook: ExcelScript.Workbook, mainSheet: ExcelScript.Worksheet) {
  const headersTable: ExcelScript.Table = workbook.getTable("Header");
  const headerSheet = workbook.getWorksheet("Header");
  const usedRange = headerSheet.getUsedRange();
  const address = usedRange.getAddress();
  const headerRange = address.split("!")[1];
  const headerObj: { [key: string]: string | string[] } = tableToObject(headersTable);

  for (let key in headerObj) {
    headerObj[key] = headerObj[key][0];
  }

  const A1 = mainSheet.getRange('A1');
  const A2 = mainSheet.getRange('A2');
  const B1 = mainSheet.getRange('B1');
  const E2 = mainSheet.getRange('E2');

  const B1D1 = mainSheet.getRange('B1:D1');
  const E2H3 = mainSheet.getRange('E2:H3');
  const E1H1 = mainSheet.getRange('E1:H1');



  const cellsToMerge = [B1D1, E2H3, E1H1];
  cellsToMerge.forEach(cell => cell.merge());

  // Setting updated by and date to non-null values
  const updatedDate = (headerObj["Created Date"] == headerObj["Modified Date"]) && !headerObj["Modified Date"] ? headerObj["Created Date"] : headerObj["Modified Date"];
  const updatedBy = (headerObj["Created By User"] == headerObj["Modified By User"]) && !headerObj["Modified By User"] ? headerObj["Created By User"] : headerObj["Modified By User"];

  // For A1 and A2 cells
  A1.setValue(headerObj["NC Number"]);
  A1.getFormat().setColumnWidth(120);


  // For other cells
  B1.setValue(headerObj["Status"]);
  E2.setValue(`Last updated on ${updatedDate.split("+")[0]}
by ${updatedBy}`);
  E2.getFormat().getFont().setColor("#6B6B6B");
  E1H1.setValue(`Assignee: ${headerObj["NC Assignee"]}`);


  A1.getFormat().getFont().setSize(16);
  A1.getFormat().getFont().setBold(true);
  B1.getFormat().getFont().setBold(true);
  B1.getFormat().getFont().setColor("White");
  B1.getFormat().getFont().setSize(13);

  // Set background color based on status value
  if (headerObj["Status"] === 'New') B1.getFormat().getFill().setColor("#1F3864");
  if (headerObj["Status"] === 'In Progress') B1.getFormat().getFill().setColor("#C65911");
  if (headerObj["Status"] === 'Closed') B1.getFormat().getFill().setColor("#71AF84");
  if (headerObj["Status"] === 'Pending Closure') B1.getFormat().getFill().setColor("#AF1909");

  const cellsToFormat: ExcelScript.Range[] = [B1D1, E2H3, E1H1];
  cellsToFormat.forEach(cells => {
    cells.getFormat().autofitColumns();
  });
  E1H1.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  E1H1.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);

  E2H3.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  B1D1.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  B1D1.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center)
}

