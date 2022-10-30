/**
 * This script allows to create a "MagicSlides" template for scalable creation
 * of presentations.
 * 
 * Prerequisites:
 *  - A Google Spreadsheet that acts as a template and is named "[TEMPLATE]..."
 *  - A Google Slide that acts as a template and is named "[TEMPLATE]..."
 */

class MagicSlides {
  // Regex to spot different kinds of substitutions
  static get regex() {
    return {
      placeholder: /\<\%.*?\%\>/gm,
      chart: /\<\%chart.*?\%\>/gm,
      table: /\<\%table.*?\%\>/gm,
      template: /.*\[TEMPLATE\]/gm
    }
  }

  static get table_chart_type_mapping() {
    return {
      string: Charts.ColumnType.STRING,
      number: Charts.ColumnType.NUMBER,
      object: Charts.ColumnType.DATE,
      default: Charts.ColumnType.STRING
    }
  }

  static get slides_table_vertival_alignment_maping() {
    return {
      top:SlidesApp.ContentAlignment.TOP,
      middle:SlidesApp.ContentAlignment.MIDDLE,
      bottom:SlidesApp.ContentAlignment.BOTTOM,
      default: SlidesApp.ContentAlignment.MIDDLE
    }
  }
  /**
   * Take in a 2d array with two columns and make an object out of it
   * 
   * @param {*[][]} values 2d array with values
   * @param {boolean} has_header flag showing whether the 1st row is a header
   * @returns {Object} Object with values in the 1st columns as keys and 2nd column as values
   */
  static valuesToDictionary(values,has_header=false){
    if(has_header){
      values.shift();
    }

    if(values[0].length > 2){
      throw Error(`Input should be a 2-dimensional array`)
    }

    return values.reduce((dictionary, row) => {
      let [key, value] = row;
      dictionary[key] = value;
      return dictionary;
    },{})
  }

  /**
   * Creates a MagicSlides object
   * 
   * @param {Object} config Information about the template that is being generated
   * @param {string} config.template_name Name of the template that is being created
   * @param {string} config.sheet_template_url URL to the template Google Sheet
   * @param {string} config.presentation_template_url URL to the template Google Slides
   * @param {string} [config.test_sheet_template_url=null] [Optional] URL to the copy of a 
   *  Google Sheets template that will be used for testing 
   * @param {string} [config.test_presentation_template_url=null] [Optional] URL to the copy of a 
   *  Google Slides template that will be used for testing 
   * @param {boolean} [config.move_templates=true] [Optional] Whether templates should be moved to a 
   * template folder (if not already there)
   * 
   * @returns {Object} - MagicSlides object
   */
  constructor({
    template_name,
    sheet_template_url, 
    presentation_template_url,
    test_sheet_template_url = null,
    test_presentation_template_url = null,
    test = false,
    move_templates = true
    }) {
    this.base_drive_folder_name = "MagicSlides";
    this.template_name = template_name;
    
    // Create (or retrieve) folder where all the materials created by this script will go to
    this.base_folder = MagicSlides.createOrRetrieveDriveFolder(DriveApp,this.base_drive_folder_name);

    // Create (or retrieve) a folder that will house the collateral
    this.collateral_folder = MagicSlides.createOrRetrieveDriveFolder(this.base_folder,this.template_name);

    // Create (or retrieve) a folder where the templates will be stored
    this.template_folder = MagicSlides.createOrRetrieveDriveFolder(this.collateral_folder,"#TEMPLATE");


    this.test = test,

    // Define templates 
    this.template_spreadsheet = SpreadsheetApp.openByUrl(sheet_template_url);
    this.template_presentation = SlidesApp.openByUrl(presentation_template_url);
    
    if(move_templates) {
      MagicSlides.moveFileToFolder(this.template_spreadsheet.getId(), this.template_folder);
      MagicSlides.moveFileToFolder(this.template_presentation.getId(), this.template_folder);
    }

    
    // Define a test sheet (i.e. a pre-made copy of the template)
    if(test_sheet_template_url) {
      this.test_template_spreadsheet = SpreadsheetApp.openByUrl(test_sheet_template_url);
    }

    // Define a test presentation (i.e. a pre-made copy of the template)
    if(test_presentation_template_url) {
      this.test_template_presentation = SlidesApp.openByUrl(test_presentation_template_url);
    }
  }

  /**
   * Check whether a file is already in a folder, and move into the folder if not
   * 
   * @param {string} file_id Drive id of the file that should be moved
   * @param {DriveApp.Folder} folder Drive folder where the file should be moved
   */
  static moveFileToFolder(file_id,folder) {
    const file = DriveApp.getFileById(file_id);
    const parents = file.getParents();

    // Check whether any of the parents are the same as the folder
    while(parents.hasNext()) {
      const parent = parents.next()

      // If there is a match, terminate the function
      if(parent.getId() === folder.getId()){
        return 
      }
    }

    // If the function didn't terminate, move the file
    file.moveTo(folder);
  }

  /**
   * Create a new Drive folder or retrieve an existing one.
   * 
   * @param {DriveApp.Folder} top_level_folder Drive (Folder) where to search for the folder
   * @param {string} folder_name Name of the folder to be retreived / created
   * @returns {DriveApp.Folder} drive_folder Drive folder for the templates
   */
  static createOrRetrieveDriveFolder(top_level_folder,folder_name){
    const folders = top_level_folder.getFoldersByName(folder_name);

    // Get all folders with the give name and push them into an array
    const drive_folders = [];
    while(folders.hasNext()) {
      drive_folders.push(folders.next());
    }

    // If there are more than one folder, 
    if(drive_folders.length > 1) {
      throw Error(`
        Multiple folders have the name '${folder_name}'.
        Please delete one of the following folders:
        ${drive_folders.map(f => f.getUrl()).join("\n")}
      `)
    }

    top_level_folder

    // If the array is empty, create a new folder called 'MagicSlides'
    if(drive_folders.length === 0) {
      const new_folder = DriveApp.createFolder(folder_name);

      // Checking whether a DriveApp or DriveApp.Folder is passed (DriveApp doesn't have getParents method)
      if(top_level_folder.getParents) {
        new_folder.moveTo(top_level_folder);
      }

      drive_folders.push(new_folder);
    }

    // Return the first item (should be the only) 
    return drive_folders[0];
  }

  /**
   * Copy the templates for the collateral
   * 
   * @param {Object} static_data Data object that contains values that the user wants to
   *    pass into the template (e.g. current date, script initiator...)
   */
  createCollateral(static_data) {
    // If the test flag is set, return the test_templates
    if(this.test) {
      return new Collateral({
        spreadsheet: this.test_template_spreadsheet,
        presentation: this.test_template_presentation,
        static_data
      })
    }

    let new_sheet_name = this.template_spreadsheet.getName().replace(MagicSlides.regex.template,"").trim();
    let new_presentation_name = this.template_presentation.getName().replace(MagicSlides.regex.template,"").trim();

    let time_string = new Date().toUTCString();  

    let new_title = static_data["title"];

    if(new_title) {
      new_title = `${new_title} (${time_string})`
    } else {
      new_title = `${time_string}`
    }

    // Create a Drive folder where the collateral will be housed
    const folder = MagicSlides.createOrRetrieveDriveFolder(this.collateral_folder,new_title);

    // Create a copy of the Sheet template and move it to the folder
    const spreadsheet = this.template_spreadsheet.copy(`${new_sheet_name} | ${new_title}`);
    MagicSlides.moveFileToFolder(spreadsheet.getId(),folder);

    // Create a copy of the Presentation template and move it to the folder
    const presentation_file = DriveApp.getFileById(this.template_presentation.getId()).makeCopy(`${new_presentation_name} | ${new_title}`, folder)
    MagicSlides.moveFileToFolder(presentation_file.getId(),folder);
    const presentation = SlidesApp.openById(presentation_file.getId());

    static_data["sheet_template_url"] = this.template_spreadsheet.getUrl();
    static_data["presentation_template_url"] = this.template_presentation.getUrl();
    static_data["sheet_url"] = spreadsheet.getUrl();
    static_data["presentation_url"] = presentation.getUrl();
    static_data["base_drive_folder_id"] = this.base_folder.getId();
    static_data["template_folder_id"] = this.template_folder.getId();
    static_data["collateral_folder_id"] = this.collateral_folder.getId();
    static_data["template_name"] = this.template_name;
    static_data["timestamp"] = new Date().toISOString();

    return new Collateral({
      spreadsheet,
      presentation,
      folder,
      static_data
    })
  }
}
