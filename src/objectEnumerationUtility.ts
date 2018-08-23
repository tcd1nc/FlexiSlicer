module powerbi.extensibility.visual {

    interface getValueParams<T> {
        maxValue?: T,
        minValue?: T
    }

    /**
     * Gets property value for a particular object.
     *
     * @function
     * @param {DataViewObjects} objects - Map of defined objects.
     * @param {string} objectName       - Name of desired object.
     * @param {string} propertyName     - Name of desired property.
     * @param {T} defaultValue          - Default value of desired property.
     */
    /* export function getValue<T>(objects: DataViewObjects, objectName: string, propertyName: string, defaultValue: T, params?: getValueParams<T>): T {
        if (objects) {
            let object = objects[objectName];
            if (object) {
                let property: T = <T>object[propertyName];
                if (property !== undefined) {
                    if (params) {
                        if (params.maxValue && property > params.maxValue) {
                            return params.maxValue;
                        }
                        if (params.minValue && property < params.minValue) {
                            return params.minValue;
                        }
                    }
                    return property;
                }
            }
        }
        return defaultValue;
    } */

    export function getValue<T>(objects: DataViewObjects, objectName: string, propertyName: string, defaultValue: T): T {
        if (objects) {
            let object = objects[objectName];
            if (object) {
                let property: T = <T>object[propertyName];
                if (property !== undefined) {
                    return property;
                }
            }
        }
        return defaultValue;
    }
    
    export function getFill(dataView: DataView, objectName: string, propertyName: string, defaultValue: string): string {
        if (dataView) {
            let objects = dataView.metadata.objects;
            if (objects) {
                var config = objects[objectName];
                if (config) {
                    let fill: Fill = <Fill>config[propertyName];
                    if (fill !== undefined && fill.solid !== undefined && fill.solid.color !== undefined)
                        return fill.solid.color;
                }
            }
        }
        //return { solid: { color: defaultValue}}; // { solid: { color: '#30ADFF' } };
        return defaultValue;
    }

    /**
     * Gets property value for a particular object in a category.
     *
     * @function
     * @param {DataViewCategoryColumn} category - List of category objects.
     * @param {number} index                    - Index of category object.
     * @param {string} objectName               - Name of desired object.
     * @param {string} propertyName             - Name of desired property.
     * @param {T} defaultValue                  - Default value of desired property.
     */
    export function getCategoricalObjectValue<T>(category: DataViewCategoryColumn, index: number, objectName: string, propertyName: string, defaultValue: T): T {
        let categoryObjects = category.objects;

        if (categoryObjects) {
            let categoryObject: DataViewObject = categoryObjects[index];
            if (categoryObject) {
                let object = categoryObject[objectName];
                if (object) {
                    let property: T = <T>object[propertyName];

                    if (property !== undefined) {
                        return property;
                    }
                }
            }
        }
        return defaultValue;
    }
}