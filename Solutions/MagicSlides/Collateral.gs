class Collateral {
  static loadCollateralFromStaticData() {
    const spreadsheet = SpreadsheetApp.getActive();
    const static_data_sheet = spreadsheet.getSheetByName("static_data");
    const static_data = MagicSlides.valuesToDictionary(
      static_data_sheet
        .getRange(1,1,static_data_sheet.getLastRow(),2)
        .getValues()
    );

    const presentation = SlidesApp.openByUrl(static_data.presentation_url);
    const folder = DriveApp.getFolderById(static_data.collateral_folder_id);

    return new Collateral({spreadsheet,presentation,folder,static_data});
  } 

  /**
   * Creates assets needed for the collataeral.
   * 
   * @param {Object} args Arguments for collaterals
   * @param {SpreadsheetApp.Spreadsheet} args.spreadsheet Spreadsheet for the collateral
   * @param {SlidesApp.Presentation} args.presentation Presentation for the collateral
   * @param {Object} args.static_data Data that the user wants to pass into each sheet
   * @param {DriveApp.Folder} args.folder Folder where the collateral is
   */
  constructor({spreadsheet,presentation,folder, static_data}) {
    this.spreadsheet = spreadsheet;
    this.presentation = presentation;
    this.folder = folder;

    this.substitutions = {};

    // Keeping track of substitutions that weren't found in all assets
    // TODO: enable tracking of substitutions that were missing from the presentation
    this.missing_substitutions = {
      presentation:[],
      sheet:[]
    };

    // Any other values passed into MagicSlides will be stored in the 'static_data' object
    this.static_data = static_data;
    this.updateStaticDataSheet();

    Logger.log(`New set of collaterals is created:\n\tSheet: ${spreadsheet.getUrl()}\n\tPresentation: ${presentation.getUrl()}`)
  }

  updateStaticDataSheet() {
    const static_data_sheet = this.getOrCreateSheet(this.spreadsheet,("static_data"));
    const current_values = static_data_sheet
      .getRange(
        1,
        1,
        static_data_sheet.getLastRow() || 1,
        2)
      .getValues()
      .reduce((acc,cur) => {
        let [key, value] = cur;

        // Create an object if the key field is populated AND if the same key doesn't
        // exist in the static values that were passed (as those will be appended, and would
        // cause duplication otherwise)
        if(key && !static_data[key]) {
          acc[key] = value;
        }
        return acc
      },{});
    
    const new_table = [
      ...Object.entries(current_values),
      ...Object.entries(this.static_data)
      ]
    
    static_data_sheet
      .getRange(1,1,new_table.length,new_table[0].length)
      .setValues(new_table);
  }

  updateSubstitutions() {

    // Step 1. Get all text substitutions
    const substitutions = {};
    const sub_sheet = this.spreadsheet.getSheetByName("substitutions");

    sub_sheet
      .getRange(1,1,sub_sheet.getLastRow() || 1, 2)
      .getValues()
      .forEach(([key,value]) => {substitutions[key] = value});

    // Step 2. Get all chart and table substitutions
    const all_sheets = this.spreadsheet.getSheets();

    // Step 2a. Filter all the tabs with chart placeholders and get their charts
    all_sheets
      .filter(sheet => sheet.getName().match(MagicSlides.regex.chart) !== null)
      .forEach(sheet => substitutions[sheet.getName()] = sheet.getCharts()[0])

    // Step 2b. Filter all tabs with tables and get their full ranges
    all_sheets
      .filter(sheet => sheet.getName().match(MagicSlides.regex.table) !== null)
      .forEach(sheet => {
        const values = sheet
          .getRange(1,1,sheet.getLastRow() || 1, sheet.getLastColumn() || 1)
          .getValues();

        const last_row_with_values = values.reduce((acc,cur,index) => {
          if(cur.join("").length > 0) {acc = index}
          return acc;
        },0);
        substitutions[sheet.getName()] = sheet.getRange(1,1,last_row_with_values + 1, sheet.getLastColumn());
      })


    this.substitutions = substitutions;
  }

  getOrCreateSheet(spreadsheet, sheet_name) {
    return spreadsheet.getSheetByName(sheet_name) || spreadsheet.insertSheet(sheet_name);
  }

  addData(sheet_name,data,has_header=true, clear_content=true, clear_format=true) {

    // Open the sheet if exists, or create a new one with the given name
    const sheet = this.getOrCreateSheet(this.spreadsheet,sheet_name);

    // If either clearing the content or format is requested, clear either or both
    if(clear_content || clear_format) {
      sheet.clear({contentsOnly:clear_content, formatOnly:clear_format})
    }

    let start_row = 1;

    // Adjust starting row if clearing content is disabled
    if(!clear_content) {

      // If the data has a header, remove it
      if(has_header) {
        data.shift()
      }

      // Check what is the last row, and increment by one
      start_row = (sheet.getLastRow() + 1) || 1
    }

    // Write data to the sheet
    sheet
      .getRange(start_row,1,data.length || 1, data[0].length || 1)
      .setValues(data);
    
    return this;
  }

  substituteText(placeholder, text_element) {
    text_element.replaceAllText(placeholder, this.substitutions[placeholder]);
  }

  createChartTable(values,height,width) {
    const data_table = Charts.newDataTable();
    // Get the headers and use them, along with the first row of data,
    // to create columns in the data_table
    values.shift().map((header, index) => {
      
      // Get the value type 
      const value_type = typeof values[0][index];
      data_table.addColumn(

        // Map the value type back to Charts.ColumnType
        MagicSlides.table_chart_type_mapping[value_type] || 
          MagicSlides.table_chart_type_mapping.default,
        header
      )
    });

    values.forEach((row) => {
      data_table.addRow(row);
    })

    return Charts.newTableChart().setDataTable(data_table).setDimensions(height,width).build().getBlob();
  }

  /**
   * Substitutes table's placeholders with tables.
   * 
   * @param {string} placeholder Placeholder string
   * @param {SlidesApp.Slide} slide Slide where the placeholder element is
   * @param {SlidesApp.PageElement} page_element Page element to be substituted§
   */
  substituteTable(placeholder, slide, page_element) {
    const data_range = this.substitutions[placeholder];
    const values = data_range.getValues();

    // Get sum of widths for every column in the range
    let data_range_width = 0;
    for(let i = 1; i < data_range.getColumn() + data_range.getWidth(); i++) {
      data_range_width += data_range.getSheet().getColumnWidth(i);
    }

    // Get sum of height for each row in the range
    let data_range_height = 0;
    for(let i = 1; i < data_range.getRow() + data_range.getHeight(); i++) {
      data_range_height += data_range.getSheet().getRowHeight(i);
    }

    // Take the min width and height begween the current element and the table
    const width = Math.min(data_range_width,page_element.getWidth());
    const height = Math.min(data_range_height,page_element.getHeight());

    // Create the table based on
    //    1. Number of columns and rows in the data range
    //    2. Position (left and top) of the placeholder element
    //    3. Width and height calculated above
    const table = slide.insertTable(
      values.length,
      values[0].length,
      page_element.getLeft(),
      page_element.getTop(),
      width,
      height
    );

    // Iterate through each cell and set the value and styling from the Google Sheet
    values.forEach((row,row_index) => {
      const table_row = table.getRow(row_index);
      row.forEach((value,column_index) => {
        const table_cell = table_row.getCell(column_index);
        const range_cell = data_range.getCell(row_index + 1, column_index + 1);
        table_cell
          .getText()
          .setText(value)
          .getTextStyle()
            .setBold(range_cell.getTextStyle().isBold())
            .setItalic(range_cell.getTextStyle().isItalic())
            .setForegroundColor(range_cell.getTextStyle().getForegroundColor())
            .setFontFamily(range_cell.getTextStyle().getFontFamily())
            .setFontSize(range_cell.getTextStyle().getFontSize())
            .setStrikethrough(range_cell.getTextStyle().isStrikethrough())
            .setUnderline(range_cell.getTextStyle().isUnderline());

        table_cell
          .getFill().setSolidFill(range_cell.getBackground());

        const vertical_alignment = range_cell.getVerticalAlignment();
        table_cell
          .setContentAlignment(
            MagicSlides.slides_table_vertival_alignment_maping[vertical_alignment] || 
              MagicSlides.slides_table_vertival_alignment_maping.default
            );
      })
    })

    // const table = this.createChartTable(values,page_element.getHeight(),page_element.getWidth())
    // slide.insertImage(table.getBlob())

    page_element.remove();
  }
  
  /**
   * Substitutes chart's placeholders with charts.
   * 
   * @param {string} placeholder Placeholder string
   * @param {SlidesApp.Slide} slide Slide where the placeholder element is
   * @param {SlidesApp.PageElement} page_element Page element to be substituted
   */
  substituteChart(placeholder, slide, page_element) {
    const chart = slide.insertSheetsChart(this.substitutions[placeholder]);
    const bounding_width = page_element.getWidth();
    const bounding_height = page_element.getHeight();

    const chart_width = chart.getWidth();
    const chart_height = chart.getHeight();

    // If the chart is within the constraints of the bounding box, do nothing
    if(bounding_width >= chart_width && bounding_height >= chart_height) {
      return
    }
    
    // Taking the highest difference between w/h and scaling the chart down
    const scaling_ratio = Math.min(bounding_width/chart_width, bounding_height/chart_height);

    chart
      .setTop(page_element.getTop()) 
      .setLeft(page_element.getLeft()) 
      .scaleWidth(scaling_ratio)
      .scaleHeight(scaling_ratio);

  page_element.remove();
  }

  updatePresentation() {
    // Get the full list of different substitutions
    this.updateSubstitutions();

    // Get all slides
    const slides = this.presentation.getSlides();
    for (let slide of slides) {

      // Iterate over each page_element on a slide
      slide.getPageElements().forEach(page_element => {

        // Get text from each element that is of the type SHAPE
        if(page_element.getPageElementType() === SlidesApp.PageElementType.SHAPE){
          // "getRuns" separates text elements where the style has changed, which
          // allow for separately styling each individual placeholder
          const text_elements = page_element.asShape().getText().getRuns();
 
          if(!text_elements) return;

          for(let text_element of text_elements) {
            const text_string = text_element.asString().trim();
            
            if(!text_string.match(MagicSlides.regex.placeholder)) {continue}

            for(let [placeholder] of text_string.matchAll(MagicSlides.regex.placeholder)) {
              
              if(!this.substitutions[placeholder]) {
                this.missing_substitutions.sheet.push(placeholder);
                continue;
              }

              if(placeholder.match(MagicSlides.regex.chart)) {
                this.substituteChart(placeholder, slide, page_element);
                return
              }

              if(placeholder.match(MagicSlides.regex.table)) {
                this.substituteTable(placeholder, slide, page_element);
                return;
              }

              this.substituteText(placeholder, text_element);
            }
          }
        }
      })
    }
  }
}

