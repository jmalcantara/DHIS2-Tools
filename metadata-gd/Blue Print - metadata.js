/**
 * Created by jmalcantara on 26/04/17.
 */
/**
 *
 *
 * Retrieves all the rows in the active spreadsheet that contain data and creates an array, indexed by its normalized column name.
 * Mandatory columns:
 * parentcode or parentid, name, shortname, code
 * Optional columns:
 * id, openingdate, comment, latitude, longitude, contactperson, address, email, phonenumber, description
 *
 * The function requires the headers in the first row, starting with the first column (A)
 * headers = A1, B1, C1...
 * data = A2, B2, C2...
 *
 *
 * To reference the different objects use the normalized column names.
 *  Example:
 *  Parent_Code -> parentcode
 *  Parent Code -> parentCode
 *
 */
function createXML() {
    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getDataRange();
    var numRows = range.getNumRows();
    var numColumns = range.getNumColumns();
    var ObjectsDataRange = sheet.getRange(2, range.getColumn(), numRows - 1, numColumns);
    var numObjectsDataRows = ObjectsDataRange.getNumRows();

    // Create objects from rows
    var metadataObjects = getRowsData(sheet, ObjectsDataRange);

    // Iterate for every row in the array to create the XML using the XmlService
    //
    //
    //
    var url = XmlService.getNamespace('http://dhis2.org/schema/dxf/2.0');
    var xmlroot = XmlService.createElement('metaData',url);

// Start

    var xmlCategories = XmlService.createElement('categories');
    var xmlCatOptionsAll = XmlService.createElement('categoryOptions');
    var xmlCategoryCombos = XmlService.createElement('categoryCombos');
    var xmlAllOptions = XmlService.createElement('options');
    var xmlCatOptionGroups = XmlService.createElement('categoryOptionGroups');
    var xmlCatOptionGroupSets = XmlService.createElement('categoryOptionGroupSets');
    var xmlOptionSets = XmlService.createElement('optionSets');
    var xmlDataElements = XmlService.createElement('dataElements');
    var xmlDataSets = XmlService.createElement('dataSets');
    var xmlSections = XmlService.createElement('sections');
    var xmlPrograms = XmlService.createElement('programs');

    var numSectionsTotal = 0;

    var newCatOption =  new Array();
    var newCategory = new Array();
    var newCatOptionGroup = new Array();

    Logger.log(numObjectsDataRows);

    for (var i = 0; i <= (metadataObjects.length - 1); i++) {

        // Categories and Category Options

        if (metadataObjects[i].categoryName != undefined) {
            //----
            var currentCategory = metadataObjects[i].categoryName;
            var CategoryDuplicate = false;
            for(j in newCategory) {
                if(currentCategory == newCategory[j]) {
                    CategoryDuplicate = true;
                }
            }
            if (!CategoryDuplicate) {
                newCategory.push(currentCategory);
                //----
                var xmlCategory = XmlService.createElement('category')
                    .setAttribute('name', metadataObjects[i].categoryName);
                if (metadataObjects[i].categoryshortName != undefined) {
                    xmlCategory.setAttribute('shortname', metadataObjects[i].categoryshortName);
                } else
                {
                    xmlCategory.setAttribute('shortname', metadataObjects[i].categoryName.substr(0,50));
                }

                var xmlDataDimenstionType = XmlService.createElement('dataDimensionType')
                    .setText('DISAGGREGATION');
                if (metadataObjects[i].categoryDimensionType != undefined) {
                    if (metadataObjects[i].categoryDimensionType == 'Attribute' ) {
                        xmlDataDimenstionType.setText('ATTRIBUTE');
                    }
                }
                xmlCategory.addContent(xmlDataDimenstionType);

                var xmldimenstionType = XmlService.createElement('dimensionType')
                    .setText('CATEGORY');
                xmlCategory.addContent(xmldimenstionType);

                var xmlDataDimension = XmlService.createElement('dataDimension')
                    .setText('false');
                if (metadataObjects[i].useAsDataDimension != undefined) {
                    if (metadataObjects[i].useAsDataDimension == 'Yes' ) {
                        xmlDataDimension.setText('true');
                    }
                }
                xmlCategory.addContent(xmlDataDimension);



                var xmlCategoryOptions = XmlService.createElement('categoryOptions');

                // Category options
                var numCategoryOptions = 0;
                for (var a = 0; a <= (metadataObjects.length - 1); a++) {
                    /*
                     [15-07-20 07:20:54:157 CDT] catOptionName
                     [15-07-20 07:20:54:159 CDT] catOptionShortName
                     [15-07-20 07:20:54:160 CDT] catOptionCode

                     Avoid duplication of category options in the same category
                     */
                    if (metadataObjects[a].catName != undefined && metadataObjects[a].catOptionName != undefined) {
                        if (metadataObjects[a].catName == currentCategory) {
                            numCategoryOptions++;
                            var currentOption = metadataObjects[a].catOptionName
                            var duplicate = false;
                            for(j in newCatOption) {
                                if(currentOption == newCatOption[j]) {
                                    duplicate = true;
                                }
                            }
                            if(!duplicate) {
                                newCatOption.push(currentOption);

                                // Category options in current category
                                var xmlCategoryOption = XmlService.createElement('categoryOption')
                                    .setAttribute('name', metadataObjects[a].catOptionName);
                                if (metadataObjects[a].catOptionCode != undefined) {
                                    xmlCategoryOption.setAttribute('code', metadataObjects[a].catOptionCode);
                                } else
                                {
                                    xmlCategoryOption.setAttribute('code', metadataObjects[a].catOptionName.substr(0,50));
                                }
                                if (metadataObjects[a].catOptionUid != undefined) {
                                    xmlCategoryOption.setAttribute('id', metadataObjects[a].catOptionUid);
                                }
                                // All category options
                                var xmlNewCategoryOption = XmlService.createElement('categoryOption')
                                    .setAttribute('name', metadataObjects[a].catOptionName);

                                if (metadataObjects[a].catOptionShortName != undefined) {
                                    xmlNewCategoryOption.setAttribute('shortName',metadataObjects[a].catOptionShortName);
                                } else
                                {
                                    xmlNewCategoryOption.setAttribute('shortName',metadataObjects[a].catOptionName.substr(0,50));
                                }
                                if (metadataObjects[a].catOptionCode != undefined) {
                                    xmlNewCategoryOption.setAttribute('code',metadataObjects[a].catOptionCode);
                                } else
                                {
                                    xmlNewCategoryOption.setAttribute('code',metadataObjects[a].catOptionName);
                                }
                                if (metadataObjects[a].catOptionUid != undefined) {
                                    xmlNewCategoryOption.setAttribute('id',metadataObjects[a].catOptionUid);
                                }
                                // Insert here start and end Dates
                                if (metadataObjects[a].catOptionStartdate != undefined) {
                                    Logger.log(metadataObjects[a].catOptionStartdate);
                                    if (metadataObjects[a].catOptionStartdate instanceof Date) {
                                        var xmlNewCatOptionstartDate = XmlService.createElement('startDate')
                                            .setText(Utilities.formatDate(metadataObjects[a].catOptionStartdate,'GMT','yyyy-MM-dd'));
                                        xmlNewCategoryOption.addContent(xmlNewCatOptionstartDate);
                                    }
                                }
                                if (metadataObjects[a].catOptionEnddate != undefined) {
                                    if (metadataObjects[a].catOptionEnddate instanceof Date) {
                                        var xmlNewCatOptionEndDate = XmlService.createElement('endDate')
                                            .setText(Utilities.formatDate(metadataObjects[a].catOptionEnddate,'GMT','yyyy-MM-dd'));
                                        xmlNewCategoryOption.addContent(xmlNewCatOptionEndDate);
                                    }
                                }


                                // Add custom attributes
                                //
                                //  WARNING! WARNING! WARNING! WARNING! WARNING! WARNING!
                                //
                                //  Custom attributes processing code is under de development and is currently hardcoded to specific attributes
                                //  to proccess another set of custom attributes the code needs to be changed.
                                //
                                //  WARNING! WARNING! WARNING! WARNING! WARNING! WARNING!
                                //
                                // 'JPI TA01 Capacity'
                                // 'JPI TA02 Cross Cutting'
                                // 'JPI TA03 Gender'
                                // 'JPI TA04 CECAP'
                                // 'JPI TA05 FP/RH'
                                // 'JPI TA06 HIV/AIDS'
                                // 'JPI TA07 MNH'
                                // 'JPI TA08 Other Infectious Diseases'
                                // 'MCSP MC01 Global-M'
                                // 'MCSP MC02 Services'
                                // 'MCSP MC03 Special'
                                //
                                // <attributeValues>
                                //   <attributeValue lastUpdated="2015-10-09T13:58:43.290+0000" created="2015-09-23T12:22:30.337+0000">
                                //     <value>true</value>
                                //     <attribute id="Bq2BWaXCTXM" name="JPI TA02: Cross-Cutting" code="TA02" created="2015-09-07T11:50:49.406+0000" lastUpdated="2015-10-08T22:49:02.034+0000"/>
                                //   </attributeValue>

                                var xmlAttibuteValues = XmlService.createElement('attributeValues');
                                var nullAttributes = true;

                                if (metadataObjects[a].jpiTa01Capacity != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueTA01 = XmlService.createElement('attributeValue');
                                    var xmlAttValueTA01 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].jpiTa01Capacity == 'No') {
                                        xmlAttValueTA01.setText('false');
                                    }

                                    var xmlAttributeTA01 = XmlService.createElement('attribute').setAttribute('code','TA06');

                                    xmlAttributeValueTA01.addContent(xmlAttValueTA01);
                                    xmlAttributeValueTA01.addContent(xmlAttributeTA01);
                                    xmlAttibuteValues.addContent(xmlAttributeValueTA01);
                                }

                                if (metadataObjects[a].jpiTa02CrossCutting != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueTA02 = XmlService.createElement('attributeValue');

                                    var xmlAttValueTA02 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].jpiTa02CrossCutting == 'No') {
                                        xmlAttValueTA02.setText('false');
                                    }

                                    var xmlAttributeTA02 = XmlService.createElement('attribute').setAttribute('code','TA08');

                                    xmlAttributeValueTA02.addContent(xmlAttValueTA02);
                                    xmlAttributeValueTA02.addContent(xmlAttributeTA02);
                                    xmlAttibuteValues.addContent(xmlAttributeValueTA02);
                                }

                                if (metadataObjects[a].jpiTa03Gender != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueTA03 = XmlService.createElement('attributeValue');

                                    var xmlAttValueTA03 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].jpiTa03Gender == 'No') {
                                        xmlAttValueTA03.setText('false');
                                    }

                                    var xmlAttributeTA03 = XmlService.createElement('attribute').setAttribute('code','TA07');

                                    xmlAttributeValueTA03.addContent(xmlAttValueTA03);
                                    xmlAttributeValueTA03.addContent(xmlAttributeTA03);
                                    xmlAttibuteValues.addContent(xmlAttributeValueTA03);
                                }

                                if (metadataObjects[a].jpiTa04Cecap != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueTA04 = XmlService.createElement('attributeValue');

                                    var xmlAttValueTA04 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].jpiTa04Cecap == 'No') {
                                        xmlAttValueTA04.setText('false');
                                    }

                                    var xmlAttributeTA04 = XmlService.createElement('attribute').setAttribute('code','TA05');

                                    xmlAttributeValueTA04.addContent(xmlAttValueTA04);
                                    xmlAttributeValueTA04.addContent(xmlAttributeTA04);
                                    xmlAttibuteValues.addContent(xmlAttributeValueTA04);
                                }

                                if (metadataObjects[a].jpiTa05Fprh != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueTA05 = XmlService.createElement('attributeValue');

                                    var xmlAttValueTA05 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].jpiTa05Fprh == 'No') {
                                        xmlAttValueTA05.setText('false');
                                    }

                                    var xmlAttributeTA05 = XmlService.createElement('attribute').setAttribute('code','TA02');

                                    xmlAttributeValueTA05.addContent(xmlAttValueTA05);
                                    xmlAttributeValueTA05.addContent(xmlAttributeTA05);
                                    xmlAttibuteValues.addContent(xmlAttributeValueTA05);
                                }

                                if (metadataObjects[a].jpiTa06Hivaids != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueTA06 = XmlService.createElement('attributeValue');

                                    var xmlAttValueTA06 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].jpiTa06Hivaids == 'No') {
                                        xmlAttValueTA06.setText('false');
                                    }

                                    var xmlAttributeTA06 = XmlService.createElement('attribute').setAttribute('code','TA01');

                                    xmlAttributeValueTA06.addContent(xmlAttValueTA06);
                                    xmlAttributeValueTA06.addContent(xmlAttributeTA06);
                                    xmlAttibuteValues.addContent(xmlAttributeValueTA06);
                                }

                                if (metadataObjects[a].jpiTa07Mnh != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueTA07 = XmlService.createElement('attributeValue');

                                    var xmlAttValueTA07 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].jpiTa07Mnh == 'No') {
                                        xmlAttValueTA07.setText('false');
                                    }

                                    var xmlAttributeTA07 = XmlService.createElement('attribute').setAttribute('code','TA03');

                                    xmlAttributeValueTA07.addContent(xmlAttValueTA07);
                                    xmlAttributeValueTA07.addContent(xmlAttributeTA07);
                                    xmlAttibuteValues.addContent(xmlAttributeValueTA07);
                                }

                                if (metadataObjects[a].jpiTa08OtherInfectiousDiseases != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueTA08 = XmlService.createElement('attributeValue');

                                    var xmlAttValueTA08 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].jpiTa08OtherInfectiousDiseases == 'No') {
                                        xmlAttValueTA08.setText('false');
                                    }

                                    var xmlAttributeTA08 = XmlService.createElement('attribute').setAttribute('code','TA04');

                                    xmlAttributeValueTA08.addContent(xmlAttValueTA08);
                                    xmlAttributeValueTA08.addContent(xmlAttributeTA08);
                                    xmlAttibuteValues.addContent(xmlAttributeValueTA08);
                                }

                                if (metadataObjects[a].mcspMc01Globalm != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueMC01 = XmlService.createElement('attributeValue');

                                    var xmlAttValueMC01 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].mcspMc01Globalm == 'No') {
                                        xmlAttValueMC01.setText('false');
                                    }

                                    var xmlAttributeMC01 = XmlService.createElement('attribute').setAttribute('code','MC01');

                                    xmlAttributeValueMC01.addContent(xmlAttValueMC01);
                                    xmlAttributeValueMC01.addContent(xmlAttributeMC01);
                                    xmlAttibuteValues.addContent(xmlAttributeValueMC01);
                                }


                                if (metadataObjects[a].mcspMc02Services != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueMC02 = XmlService.createElement('attributeValue');

                                    var xmlAttValueMC02 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].mcspMc02Services == 'No') {
                                        xmlAttValueMC02.setText('false');
                                    }

                                    var xmlAttributeMC02 = XmlService.createElement('attribute').setAttribute('code','MC02');

                                    xmlAttributeValueMC02.addContent(xmlAttValueMC02);
                                    xmlAttributeValueMC02.addContent(xmlAttributeMC02);
                                    xmlAttibuteValues.addContent(xmlAttributeValueMC02);
                                }

                                if (metadataObjects[a].mcspMc03Special != undefined) {
                                    nullAttributes = false;
                                    var xmlAttributeValueMC03 = XmlService.createElement('attributeValue');

                                    var xmlAttValueMC03 = XmlService.createElement('value')
                                        .setText('true');
                                    if (metadataObjects[a].mcspMc03Special == 'No') {
                                        xmlAttValueMC03.setText('false');
                                    }

                                    var xmlAttributeMC03 = XmlService.createElement('attribute').setAttribute('code','MC03');

                                    xmlAttributeValueMC03.addContent(xmlAttValueMC03);
                                    xmlAttributeValueMC03.addContent(xmlAttributeMC03);
                                    xmlAttibuteValues.addContent(xmlAttributeValueMC03);
                                }

                                if (!nullAttributes) {
                                    xmlNewCategoryOption.addContent(xmlAttibuteValues);
                                    nullAttributes = true;
                                }

                                xmlCatOptionsAll.addContent(xmlNewCategoryOption);
                                xmlCategoryOptions.addContent(xmlCategoryOption);
                            }

                        }
                    }
                }
                if (numCategoryOptions > 0) {
                    xmlCategory.addContent(xmlCategoryOptions);
                }
                xmlCategories.addContent(xmlCategory);
            }
        }


        // Category Combinations
        if (metadataObjects[i].categoryCombinationName != undefined) {
            var xmlCategoryCombo = XmlService.createElement('categoryCombo')
                .setAttribute('name', metadataObjects[i].categoryCombinationName);

            var xmlDataDimenstionType = XmlService.createElement('dataDimensionType')
                .setText('DISAGGREGATION');
            if (metadataObjects[i].dimensionType != undefined) {
                if (metadataObjects[i].dimensionType == 'Attribute' ) {
                    xmlDataDimenstionType.setText('ATTRIBUTE');
                }
            }
            xmlCategoryCombo.addContent(xmlDataDimenstionType);

            var xmlSkipTotal = XmlService.createElement('skipTotal')
                .setText('false');
            if (metadataObjects[i].skipCatTotalInReports != undefined) {
                if (metadataObjects[i].skipCatTotalInReports == 'Yes' ) {
                    xmlSkipTotal.setText('true');
                }
            }
            xmlCategoryCombo.addContent(xmlSkipTotal);


            var currentCategoryCombination = metadataObjects[i].categoryCombinationName;

            var xmlCatComboCategories = XmlService.createElement('categories');

            // Avoid duplicating a category in the same combination

            for (var a = 0; a <= (metadataObjects.length - 1); a++) {
                if (metadataObjects[a].catcomboName != undefined && metadataObjects[a].categoryName != undefined) {
                    if (metadataObjects[a].catcomboName == currentCategoryCombination) {
                        var xmlCCCategory = XmlService.createElement('category')
                            .setAttribute('name', metadataObjects[a].categoryName);

                        xmlCatComboCategories.addContent(xmlCCCategory);
                    }
                }
            }

            xmlCategoryCombo.addContent(xmlCatComboCategories);
            xmlCategoryCombos.addContent(xmlCategoryCombo);
        }
        // Category option groups
        /*
         [15-11-17 10:47:20:095 PST] coGroupSetName
         [15-11-17 10:47:20:096 PST] coGroupSetDescription
         [15-11-17 10:47:20:097 PST] ccValidation
         [15-11-17 10:47:20:098 PST] useAsDataDimension
         [15-11-17 10:47:20:099 PST] dimensionType
         [15-11-17 10:47:20:099 PST] space1
         [15-11-17 10:47:20:100 PST] catoptionGroupSet
         [15-11-17 10:47:20:101 PST] catoptionGroupName
         [15-11-17 10:47:20:103 PST] catoptionGroupShortName
         [15-11-17 10:47:20:104 PST] optionalName
         [15-11-17 10:47:20:104 PST] optionalShortName
         [15-11-17 10:47:20:105 PST] checkName
         [15-11-17 10:47:20:106 PST] originalCatoptionGropupCode
         [15-11-17 10:47:20:107 PST] catoptionGropupCode
         [15-11-17 10:47:20:108 PST] verification
         [15-11-17 10:47:20:109 PST] dataDimensionType
         [15-11-17 10:47:20:110 PST] space2
         [15-11-17 10:47:20:110 PST] coGroupName
         [15-11-17 10:47:20:111 PST] catOptionName
         [15-11-17 10:47:20:112 PST] catOptionShortName
         [15-11-17 10:47:20:113 PST] originalCatOptionCode
         [15-11-17 10:47:20:114 PST] catOptionCode
         */
        //

        if (metadataObjects[i].catoptionGroupName != undefined) {
            var currentCatOptionGroup = metadataObjects[i].catoptionGroupName;
            var CatOptionGroupDuplicate = false;
            for(j in newCatOptionGroup) {
                if(currentCatOptionGroup == newCatOptionGroup[j]) {
                    CatOptionGroupDuplicate = true;
                }
            }
            if (!CatOptionGroupDuplicate) {
                newCatOptionGroup.push(currentCatOptionGroup);
                //----
                var xmlCatOptionGroup = XmlService.createElement('categoryOptionGroup')
                    .setAttribute('name', metadataObjects[i].catoptionGroupName);
                if (metadataObjects[i].catoptionGroupShortName != undefined) {
                    xmlCatOptionGroup.setAttribute('shortName', metadataObjects[i].catoptionGroupShortName);
                } else
                {
                    xmlCatOptionGroup.setAttribute('shortName', metadataObjects[i].categoryName.substr(0,50));
                }
                if (metadataObjects[i].catoptionGropupCode != undefined) {
                    xmlCatOptionGroup.setAttribute('code', metadataObjects[i].catoptionGropupCode);
                }


                var xmlDataDimenstionType = XmlService.createElement('dataDimensionType')
                    .setText('DISAGGREGATION');
                if (metadataObjects[i].dataDimensionType != undefined) {
                    if (metadataObjects[i].dataDimensionType == 'Attribute' ) {
                        xmlDataDimenstionType.setText('ATTRIBUTE');
                    }
                }
                xmlCatOptionGroup.addContent(xmlDataDimenstionType);


                var xmlCategoryOptions = XmlService.createElement('categoryOptions');

                // Category options
                var numCategoryOptions = 0;
                for (var a = 0; a <= (metadataObjects.length - 1); a++) {
                    /*
                     [15-07-20 07:20:54:157 CDT] catOptionName
                     [15-07-20 07:20:54:159 CDT] catOptionShortName
                     [15-07-20 07:20:54:160 CDT] catOptionCode

                     Avoid duplication of category options in the same group
                     */
                    if (metadataObjects[a].coGroupName != undefined && metadataObjects[a].catOptionName != undefined) {
                        if (metadataObjects[a].coGroupName == currentCatOptionGroup) {
                            numCategoryOptions++;
                            var currentOption = metadataObjects[a].catOptionName
                            var duplicate = false;
                            for(j in newCatOption) {
                                if(currentOption == newCatOption[j]) {
                                    duplicate = true;
                                }
                            }
                            if(!duplicate) {
                                newCatOption.push(currentOption);

                                // Category options in current group
                                var xmlCategoryOption = XmlService.createElement('categoryOption')
                                    .setAttribute('name', metadataObjects[a].catOptionName);
                                if (metadataObjects[a].catOptionCode != undefined) {
                                    xmlCategoryOption.setAttribute('code', metadataObjects[a].catOptionCode);
                                } else
                                {
                                    xmlCategoryOption.setAttribute('code', metadataObjects[a].catOptionName.substr(0,50));
                                }
                                if (metadataObjects[a].catOptionUid != undefined) {
                                    xmlCategoryOption.setAttribute('id', metadataObjects[a].catOptionUid);
                                }
                                xmlCategoryOptions.addContent(xmlCategoryOption);
                            }
                        }
                    }
                }
                if (numCategoryOptions > 0) {
                    xmlCatOptionGroup.addContent(xmlCategoryOptions);
                    xmlCatOptionGroups.addContent(xmlCatOptionGroup);
                }
            }
        }

        //
        // Category Option Group Set
        //

        if (metadataObjects[i].coGroupSetName != undefined && metadataObjects[i].coGroupSetDescription != undefined) {
            var xmlCatOptionGroupSet = XmlService.createElement('categoryOptionGroupSet')
                .setAttribute('name', metadataObjects[i].coGroupSetName);

            var xmlDataDimenstionType = XmlService.createElement('dataDimensionType')
                .setText('DISAGGREGATION');
            if (metadataObjects[i].dimensionType != undefined) {
                if (metadataObjects[i].dimensionType == 'Attribute' ) {
                    xmlDataDimenstionType.setText('ATTRIBUTE');
                }
            }
            xmlCatOptionGroupSet.addContent(xmlDataDimenstionType);

            var xmlDimenstionType = XmlService.createElement('dimensionType')
                .setText('CATEGORYOPTION_GROUPSET');
            xmlCatOptionGroupSet.addContent(xmlDimenstionType);

            //<description>Awards</description>
            var xmlCOGroupSetDescription = XmlService.createElement('description')
                .setText(metadataObjects[i].coGroupSetDescription);
            xmlCatOptionGroupSet.addContent(xmlCOGroupSetDescription);

            var currentCatOoptionGroupSet = metadataObjects[i].coGroupSetName;

            //categoryOptionGroups
            var xmlcategoryOptionGroups = XmlService.createElement('categoryOptionGroups');

            // Avoid duplicating a category in the same combination

            for (var a = 0; a <= (metadataObjects.length - 1); a++) {
                if (metadataObjects[a].catoptionGroupSet != undefined && metadataObjects[a].catoptionGroupName != undefined && metadataObjects[a].verification == undefined) {
                    if (metadataObjects[a].catoptionGroupSet == currentCatOoptionGroupSet) {
                        var xmlCOGroup = XmlService.createElement('categoryOptionGroup')
                            .setAttribute('name', metadataObjects[a].catoptionGroupName);
                        if (metadataObjects[i].catoptionGroupShortName != undefined) {
                            xmlCOGroup.setAttribute('shortName', metadataObjects[i].catoptionGroupShortName);
                        } else
                        {
                            xmlCOGroup.setAttribute('shortName', metadataObjects[i].categoryName.substr(0,50));
                        }

                        xmlcategoryOptionGroups.addContent(xmlCOGroup);
                    }
                }
            }

            xmlCatOptionGroupSet.addContent(xmlcategoryOptionGroups);
            xmlCatOptionGroupSets.addContent(xmlCatOptionGroupSet);
        }


        // End Category option groups and groupsets


        // Data Sets

        /*
         [15-07-21 19:30:15:967 CDT] projectprogram
         [15-07-21 19:30:15:968 CDT] t1CodeY4
         [15-07-21 19:30:15:969 CDT] t1CodeY3
         [15-07-21 19:30:15:970 CDT] t1CodeY2
         [15-07-21 19:30:15:971 CDT] t1CodeY1
         [15-07-21 19:30:15:972 CDT] name
         [15-07-21 19:30:15:973 CDT] shortName
         [15-07-21 19:30:15:974 CDT] dataSetName
         [15-07-21 19:30:15:976 CDT] dataSetShortName
         [15-07-21 19:30:15:977 CDT] codebase
         [15-07-21 19:30:15:978 CDT] dataSetCode
         [15-07-21 19:30:15:980 CDT] dataSetDescription
         [15-07-21 19:30:15:981 CDT] expiryDays
         [15-07-21 19:30:15:985 CDT] daysAfterPeriodToQualifyForTimelySubmission
         [15-07-21 19:30:15:987 CDT] frequency
         [15-07-21 19:30:15:989 CDT] combinationOfCategories
         [15-07-21 19:30:15:990 CDT] approveData
         [15-07-21 19:30:15:992 CDT] allowFuturePeriods
         [15-07-21 19:30:15:995 CDT] completeAllowedOnlyIfValidationPasses
         [15-07-21 19:30:15:996 CDT] space1
         [15-07-21 19:30:15:997 CDT] dSet
         [15-07-21 19:30:15:999 CDT] dataSetSectionName
         [15-07-21 19:30:16:000 CDT] order
         [15-07-21 19:30:16:000 CDT] space2
         [15-07-21 19:30:16:001 CDT] dataset
         [15-07-21 19:30:16:003 CDT] datasetsection
         [15-07-21 19:30:16:004 CDT] outcome
         [15-07-21 19:30:16:005 CDT] program
         [15-07-21 19:30:16:005 CDT] name
         [15-07-21 19:30:16:007 CDT] shortName
         [15-07-21 19:30:16:008 CDT] dename
         [15-07-21 19:30:16:009 CDT] deshortname
         [15-07-21 19:30:16:010 CDT] codebase
         [15-07-21 19:30:16:011 CDT] code
         [15-07-21 19:30:16:012 CDT] codeValid
         [15-07-21 19:30:16:013 CDT] description
         [15-07-21 19:30:16:015 CDT] formNamelabel
         [15-07-21 19:30:16:015 CDT] decode
         [15-07-21 19:30:16:017 CDT] domainType
         [15-07-21 19:30:16:018 CDT] valueType
         [15-07-21 19:30:16:019 CDT] typeTextnumber
         [15-07-21 19:30:16:021 CDT] aggregationOperation
         [15-07-21 19:30:16:023 CDT] zeroSignificant
         [15-07-21 19:30:16:025 CDT] categoryCombination
         [15-07-21 19:30:16:026 CDT] optionsetdatavalues
         [15-07-21 19:30:16:028 CDT] optionsetcomments
         [15-07-21 19:30:16:029 CDT] legendSet
         [15-07-21 19:30:16:030 CDT] kia
         [15-07-21 19:30:16:031 CDT] deGroup
         */
        if (metadataObjects[i].dataSetName != undefined && metadataObjects[i].frequency != undefined ) {
            var xmlDataSet = XmlService.createElement('dataSet')
                .setAttribute('name', metadataObjects[i].dataSetName);
            if (metadataObjects[i].dataSetShortName != undefined) {
                xmlDataSet.setAttribute('shortName', metadataObjects[i].dataSetShortName);
            }
            // code
            if (metadataObjects[i].dataSetCode != undefined) {
                xmlDataSet.setAttribute('code',metadataObjects[i].dataSetCode);
            }
            // uid
            if (metadataObjects[i].dataSetUid != undefined) {
                xmlDataSet.setAttribute('id',metadataObjects[i].dataSetUid);
            }

            var xmlDataSetMobile = XmlService.createElement('mobile')
                .setText('false');
            xmlDataSet.addContent(xmlDataSetMobile);

            var xmlDataSetPeriod = XmlService.createElement('periodType')
                .setText(metadataObjects[i].frequency);
            xmlDataSet.addContent(xmlDataSetPeriod);

            // description
            if (metadataObjects[i].dataSetDescription != undefined) {
                var xmlDataSetDescription = XmlService.createElement('description')
                    .setText(metadataObjects[i].dataSetDescription);
                xmlDataSet.addContent(xmlDataSetDescription);
            }
            // combinationOfCategories

            var xmlDataSetCatCombo = XmlService.createElement('categoryCombo')
                .setAttribute('name', 'default');
            if (metadataObjects[i].combinationOfCategories != undefined) {
                xmlDataSetCatCombo.setAttribute('name', metadataObjects[i].combinationOfCategories);
            }
            xmlDataSet.addContent(xmlDataSetCatCombo);

            // expiryDays
            if (metadataObjects[i].expiryDays != undefined) {
                var xmlDataSetexpiryDays = XmlService.createElement('expiryDays')
                    .setText(metadataObjects[i].expiryDays);
                xmlDataSet.addContent(xmlDataSetexpiryDays);
            }
            // timelyDays
            var xmlDataSettimelyDays = XmlService.createElement('timelyDays')
                .setText('15');
            if (metadataObjects[i].daysAfterPeriodToQualifyForTimelySubmission != undefined) {
                xmlDataSettimelyDays.setText(metadataObjects[i].daysAfterPeriodToQualifyForTimelySubmission);
            }
            xmlDataSet.addContent(xmlDataSettimelyDays);

            // approveData
            var xmlDataSetapproveData = XmlService.createElement('approveData')
                .setText('false');
            if (metadataObjects[i].daysAfterPeriodToQualifyForTimelySubmission != undefined) {
                if (metadataObjects[i].daysAfterPeriodToQualifyForTimelySubmission == 'Yes') {
                    xmlDataSetapproveData.setText('true');
                }
            }
            xmlDataSet.addContent(xmlDataSetapproveData);
            // validCompleteOnly
            var xmlDataSetvalidCompleteOnly = XmlService.createElement('approveData')
                .setText('false');
            if (metadataObjects[i].completeAllowedOnlyIfValidationPasses != undefined) {
                if (metadataObjects[i].completeAllowedOnlyIfValidationPasses == 'Yes') {
                    xmlDataSetvalidCompleteOnly.setText('true');
                }
            }
            xmlDataSet.addContent(xmlDataSetvalidCompleteOnly);
            // allowFuturePeriods
            var xmlDataSetAllowFuturePeriods = XmlService.createElement('allowFuturePeriods')
                .setText('false');
            if (metadataObjects[i].allowFuturePeriods != undefined) {
                if (metadataObjects[i].allowFuturePeriods == 'Yes') {
                    xmlDataSetAllowFuturePeriods.setText('true');
                }
            }
            xmlDataSet.addContent(xmlDataSetAllowFuturePeriods);

            // Data Set Sectons

            // [15-07-21 19:30:15:997 CDT] dSet
            // [15-07-21 19:30:15:999 CDT] dataSetSectionName
            // [15-07-21 19:30:16:000 CDT] order

            var currentDataSet = metadataObjects[i].dataSetName;

            var xmlDataSetSections = XmlService.createElement('sections');
            var numSections = 0;
            try {
                for (var a = 0; a <= (metadataObjects.length - 1); a++) {
                    if (metadataObjects[a].dSet != undefined && metadataObjects[a].dataSetSectionName != undefined) {
                        if (metadataObjects[a].dSet == currentDataSet) {
                            var xmlDataSetSection = XmlService.createElement('section')
                                .setAttribute('name', metadataObjects[a].dataSetSectionName);

                            xmlDataSetSections.addContent(xmlDataSetSection);
                            numSections++;

                            // [15-07-21 19:30:16:001 CDT] dataset
                            // [15-07-21 19:30:16:003 CDT] datasetsection
                            // [15-07-21 19:30:16:008 CDT] dename

                            // Create the sections individually with the corresponding data elements

                            var xmlSection = XmlService.createElement('section')
                                .setAttribute('name', metadataObjects[a].dataSetSectionName);
                            var xmlSectionDataSet = XmlService.createElement('dataSet')
                                .setAttribute('name',currentDataSet);
                            xmlSection.addContent(xmlSectionDataSet);

                            if (metadataObjects[a].order != undefined) {
                                var xmlSectionOrder = XmlService.createElement('sortOrder')
                                    .setText(metadataObjects[a].order);
                            }

                            var xmlSectionDataElements = XmlService.createElement('dataElements');

                            var currentSection = metadataObjects[a].dataSetSectionName;
                            var numSectionDEs = 0;
                            for (var b = 0; b <= (metadataObjects.length - 1); b++) {
                                if (metadataObjects[b].dataset != undefined && metadataObjects[b].datasetsection != undefined && metadataObjects[b].dename != undefined) {
                                    if (metadataObjects[b].dataset == currentDataSet && metadataObjects[b].datasetsection == currentSection) {
                                        var xmlSectionDataElement = XmlService.createElement('dataElement')
                                            .setAttribute('name', metadataObjects[b].dename);

                                        xmlSectionDataElements.addContent(xmlSectionDataElement);
                                        numSectionDEs++;

                                    }
                                }
                            }
                            if (numSectionDEs > 0) {
                                //Add data elements to section
                                xmlSection.addContent(xmlSectionDataElements);
                            }
                            //Add section to sections
                            xmlSections.addContent(xmlSection);
                            numSectionsTotal++;
                        }
                    }
                }
            }
            catch (e) {
                Logger.log(e);
                Logger.log('Error creating data set sections')

            }

            // Data Set Sections End

            // Data Set Data Elements
            var xmlDataSetDataElements = XmlService.createElement('dataElements');
            var numDataElements = 0;
            for (var a = 0; a <= metadataObjects.length - 1; a++) {
                try {
                    if (metadataObjects[a].dataset != undefined && metadataObjects[a].dename != undefined) {
                        if (metadataObjects[a].dataset == currentDataSet) {
                            var xmlDataSetDataElement = XmlService.createElement('dataElement')
                                .setAttribute('name', metadataObjects[a].dename);

                            xmlDataSetDataElements.addContent(xmlDataSetDataElement);
                            numDataElements++;
                        }
                    }
                }
                catch (e) {
                    Logger.log(e);
                }
            }
            if (numDataElements > 0) {
                xmlDataSet.addContent(xmlDataSetDataElements);
            }
            xmlDataSets.addContent(xmlDataSet);
        }


        // Data Sets End



        // Option Sets
        // Remove spaces from code and name
        if (metadataObjects[i].optionSetName != undefined) {
            var xmlOptionSet = XmlService.createElement('optionSet')
                .setAttribute('name', metadataObjects[i].optionSetName);

            var currentOptionSet = metadataObjects[i].optionSetName;

            var xmlOptionSetOptions = XmlService.createElement('options');
            for (var a = 0; a <= (metadataObjects.length - 1); a++) {
                if (metadataObjects[a].optionSet != undefined && metadataObjects[a].option != undefined && metadataObjects[a].code != undefined && metadataObjects[a].optionSet == currentOptionSet) {
                    var optionA = XmlService.createElement('option')
                        .setAttribute('name', metadataObjects[a].option)
                        .setAttribute('code', metadataObjects[a].code);

                    var optionB = XmlService.createElement('option')
                        .setAttribute('name', metadataObjects[a].option)
                        .setAttribute('code', metadataObjects[a].code);

                    if (metadataObjects[a].optionuid != undefined) {
                        optionA.setAttribute('id', metadataObjects[a].optionuid);
                        optionB.setAttribute('id', metadataObjects[a].optionuid);
                    }

                    xmlOptionSetOptions.addContent(optionA);
                    xmlAllOptions.addContent(optionB);
                }
            }
            xmlOptionSet.addContent(xmlOptionSetOptions);
            xmlOptionSets.addContent(xmlOptionSet);
        }

        // Option Sets End

        // Data Elements
        //
        // [15-03-30 23:49:25:904 CST] interventionArea
        // [15-03-30 23:49:25:906 CST] projectprogram
        // [15-03-30 23:49:25:907 CST] name
        // [15-03-30 23:49:25:908 CST] dename
        // [15-03-30 23:49:25:909 CST] deshortname
        // [15-03-30 23:49:25:910 CST] codebase
        // [15-03-30 23:49:25:910 CST] code
        // [15-03-30 23:49:25:913 CST] codevalidation
        // [15-03-30 23:49:25:914 CST] deDescription
        // [15-03-30 23:49:25:914 CST] formName
        // [15-03-30 23:49:25:917 CST] domainType
        // [15-03-30 23:49:25:917 CST] valueType
        // [15-03-30 23:49:25:919 CST] typeTextnumber
        // [15-03-30 23:49:25:920 CST] aggregationOperation
        // [15-03-30 23:49:25:921 CST] zeroSignificant
        // [15-03-30 23:49:25:923 CST] categoryCombination
        // [15-03-30 23:49:25:924 CST] optionsetdatavalues
        // [15-03-30 23:49:25:926 CST] coreoptionsetdatavalues
        // [15-03-30 23:49:25:927 CST] optionsetcomments
        // [15-03-30 23:49:25:928 CST] legendSet

        if (metadataObjects[i].dename != undefined && metadataObjects[i].deshortname != undefined) {
            var xmlDataElement = XmlService.createElement('dataElement')
                .setAttribute('name', metadataObjects[i].dename)
                .setAttribute('shortName', metadataObjects[i].deshortname);
            if (metadataObjects[i].decode != undefined) {
                xmlDataElement.setAttribute('code', metadataObjects[i].decode)
            }
            if (metadataObjects[i].deuid != undefined) {
                xmlDataElement.setAttribute('id', metadataObjects[i].deuid)
            }

            if (metadataObjects[i].deformname != undefined) {
                var xmlDataElementFormName = XmlService.createElement('formName')
                    .setText(metadataObjects[i].deformname);
                xmlDataElement.addContent(xmlDataElementFormName);
            }

            if (metadataObjects[i].dedescription != undefined) {
                var xmlDataElementDescription = XmlService.createElement('description')
                    .setText(metadataObjects[i].dedescription);
                xmlDataElement.addContent(xmlDataElementDescription);
            }

            var xmlDataElementDomain = XmlService.createElement('domainType')
                .setText('AGGREGATE');
            var AggregateDomain = true;
            if (metadataObjects[i].domainType != undefined) {
                if (metadataObjects[i].domainType == 'Tracker') {
                    xmlDataElementDomain.setText('TRACKER');
                    AggregateDomain = false;
                }
            }
            xmlDataElement.addContent(xmlDataElementDomain);


            var xmlzeroSignificant = XmlService.createElement('zeroIsSignificant')
                .setText(false);
            if (metadataObjects[i].zeroSignificant != undefined) {
                if (metadataObjects[i].zeroSignificant == 'Yes') {
                    xmlzeroSignificant.setText(true);
                }
            }
            xmlDataElement.addContent(xmlzeroSignificant);

            // TEXT
            // LONG_TEXT
            // LETTER
            // PHONE_NUMBER
            // EMAIL
            // BOOLEAN
            // TRUE_ONLY
            // DATE
            // DATETIME
            // NUMBER
            // UNIT_INTERVAL
            // PERCENTAGE
            // INTEGER
            // INTEGER_POSITIVE
            // INTEGER_NEGATIVE
            // INTEGER_ZERO_OR_POSITIVE
            // TRACKER_ASSOCIATE
            // OPTION_SET
            // USERNAME


            /*

             USERNAME
             FILE_RESOURCE
             COORDINATE
             */


            var xmlDataElementType = XmlService.createElement('valueType')
                .setText('INTEGER');

            if (metadataObjects[i].valueType != undefined) {
                switch (metadataObjects[i].valueType) {
                    case 'Text':
                        xmlDataElementType.setText('TEXT');
                        break;
                    case 'Long text':
                        xmlDataElementType.setText('LONG_TEXT');
                        break;
                    case 'Yes Only':
                        xmlDataElementType.setText('TRUE_ONLY');
                        break;
                    case 'Yes/No':
                        xmlDataElementType.setText('BOOLEAN');
                        break;
                    case 'Date':
                        xmlDataElementType.setText('DATE');
                        break;
                    case 'User name':
                        xmlDataElementType.setText('USERNAME');
                        break;
                    case 'Unit interval':
                        xmlDataElementType.setText('UNIT_INTERVAL');
                        break;
                    case 'Number':
                        xmlDataElementType.setText('NUMBER');
                        break;
                    case 'Integer':
                        xmlDataElementType.setText('INTEGER');
                        break;
                    case 'Positive Integer':
                        xmlDataElementType.setText('INTEGER_POSITIVE');
                        break;
                    case 'Negative Integer':
                        xmlDataElementType.setText('INTEGER_NEGATIVE');
                        break;
                    case 'Positive or Zero Integer':
                        xmlDataElementType.setText('INTEGER_ZERO_OR_POSITIVE');
                        break;
                    case 'Tracker associate':
                        xmlDataElementType.setText('TRACKER_ASSOCIATE');
                        break;
                    case 'Option set':
                        xmlDataElementType.setText('OPTION_SET');
                        break;
                    case 'User name':
                        xmlDataElementType.setText('USERNAME');
                        break;
                    default:
                        xmlDataElementType.setText('INTEGER');
                        break;
                }

            }
            xmlDataElement.addContent(xmlDataElementType);
            //xmlDataElement.addContent(xmlValueType);

            var xmlAggregationOperator = XmlService.createElement('aggregationType')
                .setText('SUM')
            if (metadataObjects[i].aggregationType != undefined) {
                switch (metadataObjects[i].aggregationType) {
                    case 'Average (sum in org unit hierarchy)':
                        xmlAggregationOperator.setText('AVERAGE_SUM_ORG_UNIT');
                        break;
                    case 'Average':
                        xmlAggregationOperator.setText('AVERAGE');
                        break;
                    case 'Count':
                        xmlAggregationOperator.setText('COUNT');
                        break;
                    case 'Standard deviation':
                        xmlAggregationOperator.setText('STDDEV');
                        break;
                    case 'Variance':
                        xmlAggregationOperator.setText('VARIANCE');
                        break;
                    case 'Min':
                        xmlAggregationOperator.setText('MIN');
                        break;
                    case 'Max':
                        xmlAggregationOperator.setText('MAX');
                        break;
                    case 'None':
                        xmlAggregationOperator.setText('NONE');
                        break;
                    case 'Default':
                        xmlAggregationOperator.setText('DEFAULT');
                        break;
                    case 'Custom':
                        xmlAggregationOperator.setText('CUSTOM');
                        break;
                    default:
                        break;
                }
            }
            xmlDataElement.addContent(xmlAggregationOperator);

            var xmlCategoryCombo = XmlService.createElement('categoryCombo')
                .setAttribute('name', 'default');
            if (metadataObjects[i].coreCategoryCombination != undefined && AggregateDomain) {
                /*
                 REQUIRES THE USE OF THE categoryCombo UID? what about a code or the name?
                 */
                var xmlCategoryCombo = XmlService.createElement('categoryCombo')
                    .setAttribute('name', metadataObjects[i].coreCategoryCombination);
            }
            else {
                if (metadataObjects[i].categoryCombination != undefined && AggregateDomain) {
                    /*
                     REQUIRES THE USE OF THE categoryCombo UID? what about a code or the name?
                     */
                    var xmlCategoryCombo = XmlService.createElement('categoryCombo')
                        .setAttribute('name', metadataObjects[i].categoryCombination);
                }
            }
            xmlDataElement.addContent(xmlCategoryCombo);

            if (metadataObjects[i].optionsetdatavalues != undefined) {
                /*
                 REQUIRES THE USE OF THE optionSet UID? what about a code or the name?
                 */
                var xmlOptionSetDataValues = XmlService.createElement('optionSet')
                    .setAttribute('name', metadataObjects[i].optionsetdatavalues);
                if (metadataObjects[i].optionsetdatavaluesuid != undefined) {
                    xmlOptionSetDataValues.setAttribute('id', metadataObjects[i].optionsetdatavaluesuid);
                }
                xmlDataElement.addContent(xmlOptionSetDataValues);
            }


            xmlDataElements.addContent(xmlDataElement);
        }

        // Program
        /*
         [16-07-04 13:58:34:286 CDT] projectprogram
         [16-07-04 13:58:34:287 CDT] name
         [16-07-04 13:58:34:288 CDT] shortName
         [16-07-04 13:58:34:289 CDT] programname
         [16-07-04 13:58:34:291 CDT] programshortname
         [16-07-04 13:58:34:292 CDT] description
         [16-07-04 13:58:34:293 CDT] programtype
         [16-07-04 13:58:34:294 CDT] useRadioButtons
         [16-07-04 13:58:34:295 CDT] allowFutureEnfollmentDates
         [16-07-04 13:58:34:297 CDT] allowFutureIncidenceDates
         [16-07-04 13:58:34:299 CDT] onlyEnrollOnce
         [16-07-04 13:58:34:300 CDT] showIncidenceDate
         [16-07-04 13:58:34:301 CDT] incidentdatelabel
         [16-07-04 13:58:34:303 CDT] enrollmentdatelabel
         [16-07-04 13:58:34:304 CDT] skipOverdueEvents
         [16-07-04 13:58:34:306 CDT] programuid
         [16-07-04 13:58:34:306 CDT] spacep1
         [16-07-04 13:58:34:307 CDT] attributesprogram
         [16-07-04 13:58:34:308 CDT] attributes
         [16-07-04 13:58:34:310 CDT] displayAttributeInList
         [16-07-04 13:58:34:311 CDT] attributeIsMandatory
         [16-07-04 13:58:34:311 CDT] spacep2
         [16-07-04 13:58:34:312 CDT] stageprogram
         [16-07-04 13:58:34:314 CDT] programstagename
         [16-07-04 13:58:34:314 CDT] description
         [16-07-04 13:58:34:316 CDT] scheduleDaysFromStart
         [16-07-04 13:58:34:317 CDT] repeatable
         [16-07-04 13:58:34:319 CDT] displayGenerateEventBoxAfterCompleted
         [16-07-04 13:58:34:320 CDT] standardIntervalDays
         [16-07-04 13:58:34:322 CDT] autogenerateEvent
         [16-07-04 13:58:34:323 CDT] blockEntryFormAfterCompleted
         [16-07-04 13:58:34:325 CDT] generateEventsBasedOnEnrollmentDateNotIncidentDate
         [16-07-04 13:58:34:327 CDT] captureCoordinates
         [16-07-04 13:58:34:328 CDT] completeAllowsOnlyIfValidationPasses
         [16-07-04 13:58:34:330 CDT] pregenerateEventUid
         [16-07-04 13:58:34:331 CDT] descriptionOfReportDate
         [16-07-04 13:58:34:332 CDT] stageuid
         [16-07-04 13:58:34:333 CDT] notes
         [16-07-04 13:58:34:334 CDT] spacep3
         [16-07-04 13:58:34:335 CDT] stage
         [16-07-04 13:58:34:336 CDT] dataElements
         [16-07-04 13:58:34:337 CDT] compulsory
         [16-07-04 13:58:34:338 CDT] providedElswhere
         [16-07-04 13:58:34:339 CDT] displayInReports
         [16-07-04 13:58:34:340 CDT] dateInFuture
         */
        //
        if (metadataObjects[i].programname != undefined && metadataObjects[i].programshortname != undefined) {

            // <program name="0 ZW HTS ST - Self Test Kit Distribution" shortName="ZW HTS ST - Self Test Kit Distribution" id="M38KvTtX0nq">
            var xmlProgram = XmlService.createElement('program')
                .setAttribute('name', metadataObjects[i].programname)
                .setAttribute('shortName', metadataObjects[i].programshortname);

            if (metadataObjects[i].programuid != undefined) {
                xmlProgram.setAttribute('id', metadataObjects[i].programuid)
            }

            var xmlProgramType = XmlService.createElement('programType')
                .setText('WITHOUT_REGISTRATION');
            var xmlWithoutRegistration = XmlService.createElement('withoutRegistration')
                .setText('true');
            var xmlRegistration = XmlService.createElement('registration')
                .setText('false');
            if (metadataObjects[i].programtype != undefined && metadataObjects[i].programtype == 'MEwR') {
                xmlProgramType.setText('WITH_REGISTRATION');
                xmlWithoutRegistration.setText('false');
                xmlRegistration.setText('true');
            }
            xmlProgram.addContent(xmlProgramType);
            xmlProgram.addContent(xmlWithoutRegistration);
            xmlProgram.addContent(xmlRegistration);

            var xmlDataEntryMethod = XmlService.createElement('dataEntryMethod')
                .setText('false');
            if (metadataObjects[i].programtype != undefined && metadataObjects[i].useradiobuttons == 'Yes') {
                xmlDataEntryMethod.setText('true');
            }
            xmlProgram.addContent(xmlDataEntryMethod);

            var xmlProgramCategoryCombo = XmlService.createElement('categoryCombo')
                .setAttribute('name','default');
            if (metadataObjects[i].programcategorycombo != undefined) {
                xmlProgramCategoryCombo.setAttribute('name',metadataObjects[i].programcategorycombo);
            }
            xmlProgram.addContent(xmlProgramCategoryCombo);

            var xmlSkipOffline = XmlService.createElement('skipOffline')
                .setText('false');
            if (metadataObjects[i].skipoffline != undefined && metadataObjects[i].skipoffline == 'Yes') {
                xmlSkipOffline.setText('true');
            }
            xmlProgram.addContent(xmlSkipOffline);

            var xmldisplayFrontPageList = XmlService.createElement('displayFrontPageList')
                .setText('false');
            if (metadataObjects[i].skipoffline != undefined && metadataObjects[i].frontpagelist == 'Yes') {
                xmldisplayFrontPageList.setText('true');
            }
            xmlProgram.addContent(xmldisplayFrontPageList);


            xmlPrograms.addContent(xmlProgram);
        }


        // Program End

    }
    if (metadataObjects[0].catcomboName != undefined) {
        xmlroot.addContent(xmlCatOptionsAll);
        xmlroot.addContent(xmlCategories);
        xmlroot.addContent(xmlCategoryCombos);
    }
    if (metadataObjects[0].optionSetName != undefined) {
        xmlroot.addContent(xmlAllOptions);
        xmlroot.addContent(xmlOptionSets);
    }
    if (metadataObjects[0].catoptionGroupName != undefined) {
        xmlroot.addContent(xmlCatOptionGroups);
    }
    if (metadataObjects[0].coGroupSetName != undefined) {
        xmlroot.addContent(xmlCatOptionGroupSets);
    }

    if (metadataObjects[0].dename != undefined) {
        xmlroot.addContent(xmlDataElements);
    }
    if (metadataObjects[0].dataSetName != undefined) {
        if (numSectionsTotal > 0) {
            // Disable sections until codes are added
            //
            //xmlroot.addContent(xmlSections);
        }
        xmlroot.addContent(xmlDataSets);
    }
    if (metadataObjects[0].programname != undefined) {
        xmlroot.addContent(xmlPrograms);
    }

    var document = XmlService.createDocument(xmlroot);
    var xmldoc = XmlService.getPrettyFormat().format(document);
    Logger.log(xmldoc);

    //  Save the XML output to a document using the name of the active spreadsheet. - Verify that it's saved in the current directory -
    Logger.log(sheet.getName());
    var folder = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
    Logger.log("folder:");
    Logger.log(folder.getName());
    DriveApp.createFile(sheet.getName() + '.xml', xmldoc, 'application/xml');
    SpreadsheetApp.getActiveSpreadsheet().toast('Document "' + sheet.getName() + '.xml' + '" was created.', 'Create XML', 5);

};



// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
    columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1 ; //the range should not include the header
    var numColumns = range.getLastColumn() - range.getColumn() + 1;
    var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
    var headers = headersRange.getValues()[0];
    Logger.log('Headers length: ' + headers.length);
    Logger.log('Headers: ' + headers);
    return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
    var objects = [];
    for (var i = 0; i < data.length; ++i) {
        var object = {};
        var hasData = false;
        for (var j = 0; j < data[i].length; ++j) {
            var cellData = data[i][j];
            if (isCellEmpty(cellData)) {
                continue;
            }
            object[keys[j]] = cellData;
            hasData = true;
        }
        if (hasData) {
            objects.push(object);
        }
    }
    return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
    var keys = [];
    for (var i = 0; i < headers.length; ++i) {
        var key = normalizeHeader(headers[i]);
        if (key.length > 0) {
            keys.push(key);
            Logger.log(key);
        }
    }
    return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
//
// Normalization returns an all lowercase name if instead of spaces a symbol, hiphen or underdash is used.
//
// Examples:
//   "First_Name" -> "fistname"
//
function normalizeHeader(header) {
    var key = "";
    var upperCase = false;
    for (var i = 0; i < header.length; ++i) {
        var letter = header[i];
        if (letter == " " && key.length > 0) {
            upperCase = true;
            continue;
        }
        if (!isAlnum(letter)) {
            continue;
        }
        if (key.length == 0 && isDigit(letter)) {
            continue; // first character must be a letter
        }
        if (upperCase) {
            upperCase = false;
            key += letter.toUpperCase();
        } else {
            key += letter.toLowerCase();
        }
    }
    return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
    return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
    return char >= 'A' && char <= 'Z' ||
        char >= 'a' && char <= 'z' ||
        isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
    return char >= '0' && char <= '9';
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
    if (data.length == 0 || data[0].length == 0) {
        return null;
    }

    var ret = [];
    for (var i = 0; i < data[0].length; ++i) {
        ret.push([]);
    }

    for (var i = 0; i < data.length; ++i) {
        for (var j = 0; j < data[i].length; ++j) {
            ret[j][i] = data[i][j];
        }
    }

    return ret;
}



/**
 * Adds custom menu to the active spreadsheet whenever the
 * spreadsheet is opened.
 */
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var entries = [{
        name : "Create XML",
        functionName : "createXML"
    }];
    spreadsheet.addMenu("Scripts", entries);
};
