using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Access.Dao;
using Microsoft.Office.Interop.Access;

namespace AccessIO {
    /// <summary>
    /// Field wrapper for <see cref="dao.Field"/> objects
    /// </summary>
    public class Field : AuxiliarObject {

        private Microsoft.Office.Interop.Access.Dao.Field daoField;
        private Dictionary<string, object> props;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="daoField"><see cref="Microsoft.Office.Interop.Access.Dao.Field"/> object</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214", Justification = "base constructor is called explicitly and there is not interaction with another member variables")]
        public Field(object daoField) {
            DaoObject = daoField;
        }

        public override object DaoObject {
            get {
                return this.daoField;
            }
            set {
                if (value != null && !(value is Microsoft.Office.Interop.Access.Dao.Field))
                    throw new ArgumentException(String.Format(AccessIO.Properties.Resources.DaoObjectIsNotAValidType, typeof(Microsoft.Office.Interop.Access.Dao.Field).Name));
                this.daoField = (Microsoft.Office.Interop.Access.Dao.Field)value;
            }
        }

        public override string ClassName {
            get { return "Field"; }
        }

        public override object this[string propertyName] {
            get {
                try {
                    return this.daoField.Properties[propertyName].Value;
                } catch (Exception) {
                    return null;
                }
            }
        }

        public override void SaveProperties(ExportObject export) {
            PropertyCollectionDao propColl = new PropertyCollectionDao(daoField, daoField.Properties);
            propColl.TryWriteProperty(export, "Attributes");
            propColl.TryWriteProperty(export, "CollatingOrder");
            propColl.TryWriteProperty(export, "Type");
            propColl.TryWriteProperty(export, "Name");
            propColl.TryWriteProperty(export, "OrdinalPosition");
            propColl.TryWriteProperty(export, "Size");
            propColl.TryWriteProperty(export, "SourceField");
            propColl.TryWriteProperty(export, "SourceTable");
            propColl.TryWriteProperty(export, "DataUpdatable");
            propColl.TryWriteProperty(export, "DefaultValue");
            propColl.TryWriteProperty(export, "ValidationRule");
            propColl.TryWriteProperty(export, "ValidationText");
            propColl.TryWriteProperty(export, "Required");
            propColl.TryWriteProperty(export, "AllowZeroLength");
            propColl.TryWriteProperty(export, "VisibleValue");
            propColl.TryWriteProperty(export, "Description");
            propColl.TryWriteProperty(export, "DecimalPlaces");
            propColl.TryWriteProperty(export, "DisplayControl");
            if (propColl.PropertyHasValue("DisplayControl")) {
                //switch (daoField.Properties["DisplayControl"].Value) {
                //    case 110:   //listbox
                //        propColl.TryWriteProperty(export, "RowSourceType");
                //        propColl.TryWriteProperty(export, "RowSource");
                //        propColl.TryWriteProperty(export, "BoundColumn");
                //        propColl.TryWriteProperty(export, "ColumnCount");
                //        propColl.TryWriteProperty(export, "ColumnHeads");
                //        propColl.TryWriteProperty(export, "ColumnWidths");
                //        break;
                //    case 111:   //dropdown list
                //        propColl.TryWriteProperty(export, "RowSourceType");
                //        propColl.TryWriteProperty(export, "RowSource");
                //        propColl.TryWriteProperty(export, "BoundColumn");
                //        propColl.TryWriteProperty(export, "ColumnCount");
                //        propColl.TryWriteProperty(export, "ColumnHeads");
                //        propColl.TryWriteProperty(export, "ColumnWidths");
                //        propColl.TryWriteProperty(export, "ListRows");
                //        propColl.TryWriteProperty(export, "ListWidth");
                //        propColl.TryWriteProperty(export, "LimitToList");
                //        break;
                //}
            }
            propColl.TryWriteProperty(export, "ColumnWidth");
            propColl.TryWriteProperty(export, "ColumnOrder");
            propColl.TryWriteProperty(export, "ColumnHidden");
            propColl.TryWriteProperty(export, "Format");
            propColl.TryWriteProperty(export, "Caption");
            propColl.TryWriteProperty(export, "UnicodeCompression");
            propColl.TryWriteProperty(export, "SmartTags");
            propColl.TryWriteProperty(export, "InputMask");
        }

        public void LoadProperties(Microsoft.Office.Interop.Access.Dao.TableDef tableDef, ImportObject import) {
            props = import.ReadProperties();
            import.ReadLine(); //Reads the 'End Field' line

            daoField.Attributes = Convert.ToInt32(props["Attributes"]);

            //CollatingOrder is read only!!

            daoField.Type = Convert.ToInt16(props["Type"]);
            daoField.Name = Convert.ToString(props["Name"]);
            daoField.OrdinalPosition = Convert.ToInt16(props["OrdinalPosition"]);
            daoField.Size = Convert.ToInt32(props["Size"]);

            //SourceField, SourceTable, DataUpdatable are read only!!

            daoField.DefaultValue = Convert.ToString(props["DefaultValue"]);
            daoField.ValidationRule = Convert.ToString(props["ValidationRule"]);
            daoField.ValidationText = Convert.ToString(props["ValidationText"]);
            daoField.Required = Convert.ToBoolean(props["Required"]);

            //AllowZeroLength property is valid only for text fields
            if (daoField.Type == (short)Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText)
                daoField.AllowZeroLength = Convert.ToBoolean(props["AllowZeroLength"]);

            //VisibleValue is read only!!

        }

        public void AddCustomProperties() {
            PropertyCollectionDao propColl = new PropertyCollectionDao(daoField, daoField.Properties);
            propColl.AddOptionalProperty(props, "Description", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText);
            propColl.AddOptionalProperty(props, "DecimalPlaces", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbInteger);
            propColl.AddOptionalProperty(props, "DisplayControl", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbInteger);
            propColl.AddOptionalProperty(props, "RowSourceType", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText);
            propColl.AddOptionalProperty(props, "RowSource", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbMemo);
            propColl.AddOptionalProperty(props, "BoundColumn", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbInteger);
            propColl.AddOptionalProperty(props, "ColumnCount", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbInteger);
            propColl.AddOptionalProperty(props, "ColumnHeads", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbBoolean);
            propColl.AddOptionalProperty(props, "ColumnWidths", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText);
            propColl.AddOptionalProperty(props, "ListRows", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbInteger);
            propColl.AddOptionalProperty(props, "ListWidth", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText);
            propColl.AddOptionalProperty(props, "LimitToList", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbBoolean);

            propColl.AddOptionalProperty(props, "ColumnWidth", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbInteger);
            propColl.AddOptionalProperty(props, "ColumnOrder", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbInteger);
            propColl.AddOptionalProperty(props, "ColumnHidden", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbBoolean);
            propColl.AddOptionalProperty(props, "Format", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText);
            propColl.AddOptionalProperty(props, "Caption", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText);
            propColl.AddOptionalProperty(props, "UnicodeCompression", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbBoolean);
            propColl.AddOptionalProperty(props, "SmartTags", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText);
            propColl.AddOptionalProperty(props, "InputMask", Microsoft.Office.Interop.Access.Dao.DataTypeEnum.dbText);
        }

    }
}
