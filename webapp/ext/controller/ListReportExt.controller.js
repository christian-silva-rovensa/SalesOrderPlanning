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

                var oData = { SalesOrder: salesOrderItem[0], 
                              SalesOrderItem: salesOrderItem[1], 
                              ExpectedAvailDate: salesOrderItem[2],
                              JustExpectedAvailDate: salesOrderItem[3],
                              StockAvailDate: salesOrderItem[4],
                              JustStockAvailDate: salesOrderItem[5],
                              LoadingDate: salesOrderItem[6],
                              JustLoadingDate: salesOrderItem[7],
                              LoadingTime: salesOrderItem[8],
                              MinDeliveryDate: salesOrderItem[9],
                              JustMinDeliveryDate: salesOrderItem[10],
                              MaxDeliveryDate: salesOrderItem[11],
                              JustMaxDeliveryDate: salesOrderItem[12] };

                aData.push(oData);

            });

            var aSalesOrderItems = [];

            var bMessageSet = false;

            aData.forEach(entity => {

                if(entity.ExpectedAvailDate === undefined){
                    //entity.ExpectedAvailDate = "0000-00-00T00:00:00";
                }else{
                    entity.ExpectedAvailDate = moment(entity.ExpectedAvailDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.StockAvailDate === undefined){
                    //entity.StockAvailDate = "0000-00-00T00:00:00";
                }else{
                    entity.StockAvailDate = moment(entity.StockAvailDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.LoadingDate === undefined){
                    //entity.LoadingDate = "0000-00-00T00:00:00";
                }else{
                    entity.LoadingDate = moment(entity.LoadingDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.LoadingTime === undefined){
                    //entity.LoadingTime = "";
                }else{
                    let text = entity.LoadingTime.toString();
                    const myArr = text.split(":");
                    if(myArr.length === 3){
                        entity.LoadingTime = "PT"+myArr[0]+"H"+myArr[1]+"M"+myArr[2]+"S";
                    }
                }

                if(entity.MinDeliveryDate === undefined){
                    //entity.MinDeliveryDate = "0000-00-00T00:00:00";
                }else{
                    entity.MinDeliveryDate = moment(entity.MinDeliveryDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                if(entity.MaxDeliveryDate === undefined){
                    //entity.MaxDeliveryDate = "0000-00-00T00:00:00";
                }else{
                    entity.MaxDeliveryDate = moment(entity.MaxDeliveryDate).format('YYYY-MM-DDTHH:MM:SS');
                }

                var oSalesOrderItem = { SalesOrder: entity.SalesOrder, 
                                        SalesOrderItem: entity.SalesOrderItem, 
                                        ExpectedAvailDate: entity.ExpectedAvailDate,
                                        StockAvailDate: entity.StockAvailDate,
                                        LoadingDate: entity.LoadingDate,
                                        LoadingTime: entity.LoadingTime,
                                        MinDeliveryDate: entity.MinDeliveryDate,
                                        MaxDeliveryDate: entity.MaxDeliveryDate,
                                        JustExpectedAvailDate: entity.JustExpectedAvailDate,
                                        JustStockAvailDate: entity.JustStockAvailDate,
                                        JustLoadingDate: entity.JustLoadingDate,
                                        JustMinDeliveryDate: entity.JustMinDeliveryDate,
                                        JustMaxDeliveryDate: entity.JustMaxDeliveryDate,
                                     };
                
                aSalesOrderItems.push(oSalesOrderItem);

            });

            if( aSalesOrderItems.length > 0 && bMessageSet === false){
                //TODO: i18n
                this._setCustomMessage("Planning is being updated", "I");
                bMessageSet = true;
            }

            aSalesOrderItems.forEach(entity => {
            
                // Bypass ETag optimistic locking when updating entities via file upload
                oModel.setHeaders({"If-Match": "*"});

                oModel.update("/Planning(SalesOrder='"+entity.SalesOrder+"',SalesOrderItem='"+entity.SalesOrderItem+"')",
                    entity, {
                    success: function(response){
                        
                        if(!response){
                            //TODO: i18n
                            let sMessage = "Sales Order: "+entity.SalesOrder + ", Sales Order Item: " + entity.SalesOrderItem + " updated successfully.";
                            that.addMessage(sMessage, "S");
                        }

                    }, 
                    error: function(err){

                        return;

                        let sMessage = "An unexpected error occurred";

                        if(!err.responseText){
                            that.addMessage(sMessage, "E");
                        }else{

                            let oErr = JSON.parse(err.responseText);

                            if (oErr.error.message.value) {
                                sMessage = oErr.error.message.value;
                            }
                            
                            that.addMessage(sMessage, "E");

                        }

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
                                 RequestFinishDate: element.getProperty("ExpectedAvailDate"),
                                 JustRequestFinishDate: element.getProperty("JustExpectedAvailDate"),
                                 FinalFinishDate: element.getProperty("StockAvailDate"),
                                 JustFinalFinishDate: element.getProperty("JustStockAvailDate"),
                                 ExpectedLoadingDate: element.getProperty("LoadingDate"),
                                 JustExpectedLoadingDate: element.getProperty("JustLoadingDate"),
                                 ExpectedLoadingTime: element.getProperty("LoadingTime"),
                                 MinDeliveryDate: element.getProperty("MinDeliveryDate"),
                                 JustMinDeliveryDate: element.getProperty("JustMinDeliveryDate"),
                                 MaxDeliveryDate: element.getProperty("MaxDeliveryDate"),                                 
                                 JustMaxDeliveryDate: element.getProperty("JustMaxDeliveryDate")
                            });

                    /*
                    oData.push({ SalesOrder: element.getProperty("SalesOrder"), 
                                 SalesOrderItem: element.getProperty("SalesOrderItem"), 
                                 DeliveryDate: element.getProperty("DeliveryDate"),
                                 Material: element.getProperty("Material"),
                                 OrderQuantity: element.getProperty("OrderQuantity"),
                                 OrderQuantityUnit: element.getProperty("OrderQuantityUnit"),
                                 Plant: element.getProperty("Plant"),
                                 RequestFinishDate: element.getProperty("ExpectedAvailDate"),
                                 JustRequestFinishDate: element.getProperty("JustExpectedAvailDate"),
                                 FinalFinishDate: element.getProperty("StockAvailDate"),
                                 JustFinalFinishDate: element.getProperty("JustStockAvailDate"),
                                 ExpectedLoadingDate: element.getProperty("LoadingDate"),
                                 JustExpectedLoadingDate: element.getProperty("JustLoadingDate"),
                                 ExpectedLoadingTime: element.getProperty("LoadingTime"),
                                 MinDeliveryDate: element.getProperty("MinDeliveryDate"),
                                 JustMinDeliveryDate: element.getProperty("JustMinDeliveryDate"),
                                 MaxDeliveryDate: element.getProperty("MaxDeliveryDate"),                                 
                                 JustMaxDeliveryDate: element.getProperty("JustMaxDeliveryDate")
                            });
                    */
                });

            }else{
                
                //Sample data
                oData.push({ SalesOrder:"9999999999", 
                             SalesOrderItem:"10", 
                             RequestFinishDate: "31/12/9999",
                             JustRequestFinishDate: "",
                             FinalFinishDate: "31/12/9999",
                             JustFinalFinishDate: "",
                             ExpectedLoadingDate: "31/12/9999",
                             JustExpectedLoadingDate: "",
                             ExpectedLoadingTime: "23:59:59",
                             MinDeliveryDate: "31/12/9999",
                             JustMinDeliveryDate: "",
                             MaxDeliveryDate: "31/12/9999",                           
                             JustMaxDeliveryDate: "" }
                        );
            }

            const worksheet = XLSX.utils.json_to_sheet(oData);

            const workbook = XLSX.utils.book_new();
            
            XLSX.utils.book_append_sheet(workbook, worksheet, "Planning");     

            XLSX.writeFile(workbook, "PlanningTemplate.xlsx", { compression: true });

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