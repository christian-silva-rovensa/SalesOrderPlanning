sap.ui.define([
    "sap/m/MessageToast",
    "sap/m/MessagePopover",
    "sap/m/MessagePopoverItem",
    "rovensa/soplan/libs/moment-with-locales.min",
    "rovensa/soplan/libs/xlsx.full.min"
], function(MessageToast, MessagePopover, MessagePopoverItem) {
    'use strict';

    return {
        FileUpload: function(oEvent) {
            
            this.onOpenDialog();         

        },

        async onOpenDialog() {

			// Create dialog lazily
			this.oDialog ??= await this.loadFragment({
				name: "rovensa.soplan.ext.fragment.UploadDialog",
                controller: this
			});

			this.oDialog.open();
		},

        onCloseDialog: function(oEvent) {
            var oFileUploader = this.byId("fileUploader");
            oFileUploader.clear();
            this.oDialog.close();
        },

        handleFileChange : function(oEvent){
            this._file = oEvent.getParameter("files") && oEvent.getParameter("files")[0];
        },

        handleFileTypeMissmatch: function(oEvent) {
            MessageToast.show("Only Excel files allowed");
        },

        handleUploadPress: function(oEvent) {

            if (!this._file) {
                return;
            }

            var oMessageManager = sap.ui.getCore().getMessageManager();

            oMessageManager.removeAllMessages();

            var that = this;
            
            var oFileUploader = this.byId("fileUploader");
			
            oFileUploader.checkFileReadable().then(async function() {
		              
                that._handleFileAsync(that.getView().getModel("custom"));

                that.oDialog.close();

			}, function(error) {
				MessageToast.show("The file cannot be read. It may have changed.");
			}).then(function() {
				oFileUploader.clear();
			});
        },

        async _handleFileAsync(oModel) {

            if (!this._file){
                return;
            }

            let that = this;
            
            const data = await this._file.arrayBuffer();
            
            // SheetJS - https://git.sheetjs.com/sheetjs/sheetjs
            const workbook = XLSX.read(data);

            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                
            const raw_data = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false});

            //console.log(raw_data);

            if (raw_data.length < 2){ 
                return;
            }

            var aData = [];

            //slice(1) is used here to skip the header
            raw_data.slice(1).forEach(salesOrderItem => {

                let oData = { SalesOrder: salesOrderItem[0], 
                              SalesOrderItem: salesOrderItem[1], 
                              ExpectedAvailDate: salesOrderItem[8],
                              JustExpectedAvailDate: salesOrderItem[9],
                              StockAvailDate: salesOrderItem[10],
                              JustStockAvailDate: salesOrderItem[11],
                              StockAvailDateNotes: salesOrderItem[12],
                              ConfirmedLoadingDate: salesOrderItem[13],
                              LoadingTime: salesOrderItem[14],
                              LoadingDate: salesOrderItem[15],
                              JustLoadingDate: salesOrderItem[16],
                              LoadingDateNotes: salesOrderItem[17],
                              VesselETD: salesOrderItem[18],
                              MinDeliveryDate: salesOrderItem[19],
                              JustMinDeliveryDate: salesOrderItem[20],
                              MaxDeliveryDate: salesOrderItem[21],
                              JustMaxDeliveryDate: salesOrderItem[22] };

                aData.push(oData);

            });

            var aSalesOrderItems = [];

            aData.forEach(entity => {

                if(entity.ExpectedAvailDate !== undefined){
                    entity.ExpectedAvailDate = moment(entity.ExpectedAvailDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.StockAvailDate !== undefined){
                    entity.StockAvailDate = moment(entity.StockAvailDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.ConfirmedLoadingDate !== undefined){
                    entity.ConfirmedLoadingDate = moment(entity.ConfirmedLoadingDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.LoadingDate !== undefined){
                    entity.LoadingDate = moment(entity.LoadingDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.LoadingTime !== undefined){
                    let str = entity.LoadingTime.toString();
                    let matches = str.replace(/[^0-9:]/g, "");
                    const myArr = matches.split(":");
                    if(myArr.length === 3){
                        entity.LoadingTime = "PT"+myArr[0]+"H"+myArr[1]+"M"+myArr[2]+"S";
                    }else if(myArr.length === 2){
                        entity.LoadingTime = "PT"+myArr[0]+"H"+myArr[1]+"M00S";
                    }else{
                        entity.LoadingTime = "PT00H00M00S";
                    }
                }

                if(entity.VesselETD !== undefined){
                    entity.VesselETD = moment(entity.VesselETD).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.MinDeliveryDate !== undefined){
                    entity.MinDeliveryDate = moment(entity.MinDeliveryDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.MaxDeliveryDate !== undefined){
                    entity.MaxDeliveryDate = moment(entity.MaxDeliveryDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                var oSalesOrderItem = { SalesOrder: entity.SalesOrder, 
                                        SalesOrderItem: entity.SalesOrderItem, 
                                        ExpectedAvailDate: entity.ExpectedAvailDate,
                                        StockAvailDate: entity.StockAvailDate,
                                        ConfirmedLoadingDate: entity.ConfirmedLoadingDate,
                                        LoadingTime: entity.LoadingTime,
                                        LoadingDate: entity.LoadingDate,
                                        VesselETD: entity.VesselETD,
                                        MinDeliveryDate: entity.MinDeliveryDate,
                                        MaxDeliveryDate: entity.MaxDeliveryDate,
                                        JustExpectedAvailDate: that._addLeadingZero(entity.JustExpectedAvailDate),
                                        JustStockAvailDate: that._addLeadingZero(entity.JustStockAvailDate),
                                        JustLoadingDate: that._addLeadingZero(entity.JustLoadingDate),
                                        JustMinDeliveryDate: that._addLeadingZero(entity.JustMinDeliveryDate),
                                        JustMaxDeliveryDate: that._addLeadingZero(entity.JustMaxDeliveryDate),
                                        StockAvailDateNotes: entity.StockAvailDateNotes,
                                        LoadingDateNotes: entity.LoadingDateNotes
                                     };
                
                aSalesOrderItems.push(oSalesOrderItem);

            });

            if( aSalesOrderItems.length > 0){
                //TODO: i18n
                //this._setCustomMessage("Planning is being updated", "I");
                MessageToast.show("Planning is being updated");
            }

            aSalesOrderItems.forEach(entity => {
            
                // Bypass ETag optimistic locking when updating entities via file upload
                oModel.setHeaders({"If-Match": "*"});

                oModel.update("/Planning(SalesOrder='"+entity.SalesOrder+"',SalesOrderItem='"+entity.SalesOrderItem+"')",
                    entity, {
                    success: function(response){
                        
                        if(!response){
                            //TODO: i18n
                            let sMessage = "Sales Order: "+ entity.SalesOrder + ", Sales Order Item: " + entity.SalesOrderItem + " updated successfully.";
                            that.addMessage(sMessage, "S");
                        }

                    }, 
                    error: function(err){
                        
                        /*
                        if(err.responseText){

                            let oError = JSON.parse(err.responseText);

                            if (oError.error.message.value) {
                                // Add a message only if the error message does not contain the Sales Ordem information  
                                if (!oError.error.message.value.includes(entity.SalesOrder)) {
                                    let sMessage = "Sales Order: "+ entity.SalesOrder + ", Sales Order Item: " + entity.SalesOrderItem + " - " + oError.error.message.value;
                                    that.addMessage(sMessage, "E");
                                }
                            }
                            
                        }
                        */

                    }
                });

            });

            that.DisplayMessages();

        },

        DisplayMessages: function(oEvent) {

            var oMP = new MessagePopover({
                items: {
                    path:"message>/",
                    template: new MessagePopoverItem({ description: "{message>description}", type: "{message>type}", title: "{message>message}"})
                }
            });

            var oMessageManager = sap.ui.getCore().getMessageManager();
            
            oMP.setModel(oMessageManager.getMessageModel(),"message");

            var oDisplayMessagesButton = this.byId("DisplayMessagesButton");

            oMP.openBy(oDisplayMessagesButton);
        },
        
        DownloadTemplate: function(oEvent) {

            let aContexts = this.extensionAPI.getSelectedContexts();

            var oData = [];

            if (aContexts.length > 0) {

                aContexts.forEach(element => {
                    oData.push({ SalesOrder: element.getProperty("SalesOrder"), 
                                 SalesOrderItem: element.getProperty("SalesOrderItem"), 
                                 Material: element.getProperty("Material"),
                                 MaterialName: element.getProperty("MaterialName"), 
                                 OrderQuantity: element.getProperty("OrderQuantity"), 
                                 OrderQuantityUnit: element.getProperty("OrderQuantityUnit"),
                                 SoldToParty: element.getProperty("SoldToParty"),
                                 SoldToPartyName: element.getProperty("SoldToPartyName"),
                                 RequestFinishDate: element.getProperty("ExpectedAvailDate"),
                                 JustRequestFinishDate: element.getProperty("JustExpectedAvailDate"),
                                 FinalFinishDate: element.getProperty("StockAvailDate"),
                                 JustFinalFinishDate: element.getProperty("JustStockAvailDate"),
                                 FinalFinishDateNotes: element.getProperty("StockAvailDateNotes"),
                                 ConfirmedLoadingDate: element.getProperty("ConfirmedLoadingDate"),
                                 ConfirmedLoadingTime: element.getProperty("LoadingTime"),
                                 ExpectedLoadingDate: element.getProperty("LoadingDate"),
                                 JustExpectedLoadingDate: element.getProperty("JustLoadingDate"),
                                 ExpectedLoadingDateNotes: element.getProperty("LoadingDateNotes"),
                                 VesselETD: element.getProperty("VesselETD"),
                                 MinDeliveryDate: element.getProperty("MinDeliveryDate"),
                                 JustMinDeliveryDate: element.getProperty("JustMinDeliveryDate"),
                                 MaxDeliveryDate: element.getProperty("MaxDeliveryDate"),                                 
                                 JustMaxDeliveryDate: element.getProperty("JustMaxDeliveryDate")
                            });

                });

            }else{
                
                //Sample data
                oData.push({ SalesOrder: "9999999999", 
                             SalesOrderItem: "10", 
                             Material: "",
                             MaterialName: "", 
                             OrderQuantity: "", 
                             OrderQuantityUnit: "",
                             SoldToParty: "",
                             SoldToPartyName: "",
                             RequestFinishDate: "31/12/9999",
                             JustRequestFinishDate: "",
                             FinalFinishDate: "31/12/9999",
                             JustFinalFinishDate: "",
                             FinalFinishDateNotes: "",
                             ConfirmedLoadingDate: "31/12/9999",
                             ConfirmedLoadingTime: "23:59:59",
                             ExpectedLoadingDate: "31/12/9999",
                             JustExpectedLoadingDate: "",
                             ExpectedLoadingDateNotes: "",
                             VesselETD: "31/12/9999",
                             MinDeliveryDate: "31/12/9999",
                             JustMinDeliveryDate: "",
                             MaxDeliveryDate: "31/12/9999",                                 
                             JustMaxDeliveryDate: ""
                    });
            }

            const worksheet = XLSX.utils.json_to_sheet(oData);

            const workbook = XLSX.utils.book_new();
            
            XLSX.utils.book_append_sheet(workbook, worksheet, "Planning");     

            XLSX.writeFile(workbook, "PlanningTemplate.xlsx", { compression: true });

        },

        _addLeadingZero: function(sValue) {

            if (sValue === undefined) {
                return;
            }

            let sValueNew = sValue; 

            const regex = new RegExp('[0-9]');

            try {
              
                if(sValueNew.length === 1 && regex.test(sValueNew) === true) {
                    sValueNew = '0'+sValueNew;
                }
              
            } catch (error) {
                console.log(error);        
            }

            return sValueNew;

        },

        _getMessageType: function(sType) {

            var sMessageType = sap.ui.core.MessageType.None;

            switch (sType) {
             
                case "S":
                    sMessageType = sap.ui.core.MessageType.Success;
                    break;
                
                case "E":
                    sMessageType = sap.ui.core.MessageType.Error;
                    break;
                
                case "W":
                    sMessageType = sap.ui.core.MessageType.Warning;
                    break;
                
                case "I":
                    sMessageType = sap.ui.core.MessageType.Information;
                    break;

            }

            return sMessageType;

        },

        _setCustomMessage: function(sMessage, sType) {

            var oMessage = {
                message: sMessage,
                type: this._getMessageType(sType)
            };

            this.extensionAPI.setCustomMessage(oMessage);

        },

        addMessage: function(sMessage, sType) {

            var oMessageManager = sap.ui.getCore().getMessageManager();

            var oMessage = new sap.ui.core.message.Message({
                message: sMessage,
                persistent: true, // create message as transition message
                type: this._getMessageType(sType)
            });

            oMessageManager.addMessages(oMessage);

        }
    };
});