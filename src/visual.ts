/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

module powerbi.extensibility.visual {
    "use strict";
    interface FlexiSlicerViewModel {
        vertical: boolean;
        defaultId: number;
        numberCats: number;
        fontSize: number;
        fontFamily: string;
        fontColor: string;
        backgroundColor: string;
    };

    function visualTransform(options: VisualUpdateOptions, host: IVisualHost): FlexiSlicerViewModel {
        let dataViews = options.dataViews;
        let viewModel: FlexiSlicerViewModel = {
           vertical: true,
           defaultId: 0,
           numberCats: 0,
           fontSize: 12,
           fontFamily: "Arial",
           fontColor: "blue",
           backgroundColor: "white"
        };
       
        if (!dataViews
            || !dataViews[0]
            || !dataViews[0].categorical
            || !dataViews[0].categorical.categories
            || !dataViews[0].categorical.categories[0]
        )
            return viewModel;        
        
        let category = options.dataViews[0].categorical.categories[0];
        var numCats = category.values.length;

        let dvobjs = dataViews[0].metadata.objects;
        try{
            let style: FlexiSlicerViewModel = {
                    vertical: getValue<boolean>(dvobjs, 'labels', 'vertical', true),
                    defaultId: getValue<number>(dvobjs, 'defaults', 'defaultId', 0),
                    numberCats: numCats,
                    fontSize: getValue<number>(dvobjs, 'labels', 'fontSize', 12),
                    fontFamily: getValue<string>(dvobjs, 'labels', 'fontFamily', "Arial"),
                    fontColor: getFill(dataViews[0], 'labels', 'fontColor', "blue"),
                    backgroundColor: getFill(dataViews[0], 'labels', 'backgroundColor', "white"),
                };
                
            return style;
        }
        catch(e){
            return viewModel;
        }
    }


    export class Visual implements IVisual {
        private target: HTMLElement;
        private flexiSlicerViewModel: FlexiSlicerViewModel;
        //private numberCats: number;
        private selectionManager: ISelectionManager;
        private selectionIds: any = {};
        private host: IVisualHost;
        private isEventUpdate: boolean = false;
        private lastSelectedValue: any;

        constructor(options: VisualConstructorOptions) {
            this.target = options.element;
            this.host = options.host;
            this.selectionManager = options.host.createSelectionManager();
        }

        public update(options: VisualUpdateOptions) {
            this.flexiSlicerViewModel = visualTransform(options, this.host);
            if (this.flexiSlicerViewModel && !this.isEventUpdate){             
                this.init(options);

                //update if default item is set
                if(this.flexiSlicerViewModel.defaultId > 0){
                    this.selectionManager.clear(); // Clean up previous filter before applying another one.                    
                    // Find the selectionId and select it
                    this.selectionManager.select(this.lastSelectedValue).then((ids: ISelectionId[]) => {  });                
                    // This call applys the previously selected selectionId                    
                    this.selectionManager.applySelectionFilter();
                }
            }
        }

        public init(options: VisualUpdateOptions) {
            // Return if we don't have a category
            if (!options ||
                !options.dataViews ||
                !options.dataViews[0] ||
                !options.dataViews[0].categorical ||
                !options.dataViews[0].categorical.categories ||
                !options.dataViews[0].categorical.categories[0]) {
                return;
            }
            let viewmodel = this.flexiSlicerViewModel;

            // remove any children from previous renders
            while (this.target.firstChild) {
                this.target.removeChild(this.target.firstChild);
            }

            // clear out any previous selection ids
            this.selectionIds = {};

            // get the category data.
            let category = options.dataViews[0].categorical.categories[0];
            let values = category.values;

            // build selection ids to be used by filtering capabilities later
            var itemctr: number = 0;
            let scroller = document.createElement("div");
            scroller.className="container";
            scroller.style.width = options.viewport.width.toString() +"px";
            scroller.style.height = options.viewport.height.toString() +"px";
            scroller.style.backgroundColor = viewmodel.backgroundColor;

            values.forEach((item: number, index: number) => {
                itemctr++;
                // create an in-memory version of the selection id so it can be used in onclick event.
                this.selectionIds[item] = this.host.createSelectionIdBuilder()
                    .withCategory(category, index)
                    .createSelectionId();               

                let value = item.toString();
                let radio = document.createElement("input");
                radio.type = "radio";
                radio.value = value;
                radio.name = "values";

                //set default checked item
                if(itemctr == viewmodel.defaultId ){
                    radio.checked = true;
                    this.lastSelectedValue = this.selectionIds[value];
                }

                radio.onclick = function (ev) {
                    this.isEventUpdate = true;
                     // This is checked in the update method. If true it won't re-render, this prevents an infinite loop                   
                    this.selectionManager.clear(); // Clean up previous filter before applying another one.                    
                    // select saved selectionid
                    this.selectionManager.select(this.selectionIds[value]).then((ids: ISelectionId[]) => {  });                
                    // This call applys the previously saved selectionId                    
                    this.selectionManager.applySelectionFilter();

                }.bind(this);
            
                let label = document.createElement("label");
                label.innerHTML += value;
                label.style.fontFamily= viewmodel.fontFamily;
                label.style.color = viewmodel.fontColor;
                label.style.fontSize = viewmodel.fontSize.toString()+"px";

                scroller.appendChild(radio);
                scroller.appendChild(label);

                if (viewmodel.vertical == true && itemctr < values.length ) {
                    scroller.appendChild(document.createElement("br"));
                }

            });
           
            this.target.appendChild(scroller);
        }


        /**
         * Enumerates through the objects defined in the capabilities and adds the properties to the format pane
         *
         * @function
         * @param {EnumerateVisualObjectInstancesOptions} options - Map of defined objects
         */
        //@logExceptions()
        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
            let objectName = options.objectName;
            let objectEnumeration: VisualObjectInstance[] = [];
            let viewModel = this.flexiSlicerViewModel;
            switch (objectName) {
                case 'defaults':
                    var id: number = 0;
                    if(viewModel.defaultId > viewModel.numberCats) id= 0; else id= viewModel.defaultId;

                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {                           
                            defaultId: id
                        },
                        validValues: {                            
                            defaultId: {
                                numberRange: {
                                    min: 0,
                                    max: viewModel.numberCats
                                }
                            }
                        },
                        selector: null
                    });
                    break;
                case 'labels':
                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {
                            vertical: viewModel.vertical,
                            backgroundColor: viewModel.backgroundColor,
                            fontSize: viewModel.fontSize,
                            fontFamily: viewModel.fontFamily,
                            fontColor: viewModel.fontColor
                        },
                        selector: null
                    });
                    break;                
            };
            this.isEventUpdate = false;
            
            return objectEnumeration;
        }
    }
}
       /*  public destroy(): void {
            //TODO: Perform any cleanup tasks here
        } */
