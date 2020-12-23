using System;

namespace WireMarking
{
    // NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
    public partial class EplanLabelling
    {

        private EplanLabellingDocument documentField;

        private decimal versionField;

        /// <remarks/>
        public EplanLabellingDocument Document
        {
            get
            {
                return this.documentField;
            }
            set
            {
                this.documentField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public decimal version
        {
            get
            {
                return this.versionField;
            }
            set
            {
                this.versionField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EplanLabellingDocument
    {

        private EplanLabellingDocumentPage pageField;

        private byte source_idField;

        /// <remarks/>
        public EplanLabellingDocumentPage Page
        {
            get
            {
                return this.pageField;
            }
            set
            {
                this.pageField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte source_id
        {
            get
            {
                return this.source_idField;
            }
            set
            {
                this.source_idField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EplanLabellingDocumentPage
    {

        private EplanLabellingDocumentPageHeader headerField;

        private EplanLabellingDocumentPageColumnHeader[] columnHeaderField;

        private EplanLabellingDocumentPageLine[] lineField;

        private EplanLabellingDocumentPageFooter footerField;

        private byte source_idField;

        /// <remarks/>
        public EplanLabellingDocumentPageHeader Header
        {
            get
            {
                return this.headerField;
            }
            set
            {
                this.headerField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ColumnHeader")]
        public EplanLabellingDocumentPageColumnHeader[] ColumnHeader
        {
            get
            {
                return this.columnHeaderField;
            }
            set
            {
                this.columnHeaderField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Line")]
        public EplanLabellingDocumentPageLine[] Line
        {
            get
            {
                return this.lineField;
            }
            set
            {
                this.lineField = value;
            }
        }

        /// <remarks/>
        public EplanLabellingDocumentPageFooter Footer
        {
            get
            {
                return this.footerField;
            }
            set
            {
                this.footerField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte source_id
        {
            get
            {
                return this.source_idField;
            }
            set
            {
                this.source_idField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EplanLabellingDocumentPageHeader
    {

        private EplanLabellingDocumentPageHeaderProperty propertyField;

        /// <remarks/>
        public EplanLabellingDocumentPageHeaderProperty Property
        {
            get
            {
                return this.propertyField;
            }
            set
            {
                this.propertyField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EplanLabellingDocumentPageHeaderProperty
    {

        private string propertyNameField;

        private string propertyValueField;

        private byte formattingTypeField;

        private byte formattingLengthField;

        private byte formattingRAlignField;

        /// <remarks/>
        public string PropertyName
        {
            get
            {
                return this.propertyNameField;
            }
            set
            {
                this.propertyNameField = value;
            }
        }

        /// <remarks/>
        public string PropertyValue
        {
            get
            {
                return this.propertyValueField;
            }
            set
            {
                this.propertyValueField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte FormattingType
        {
            get
            {
                return this.formattingTypeField;
            }
            set
            {
                this.formattingTypeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte FormattingLength
        {
            get
            {
                return this.formattingLengthField;
            }
            set
            {
                this.formattingLengthField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte FormattingRAlign
        {
            get
            {
                return this.formattingRAlignField;
            }
            set
            {
                this.formattingRAlignField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EplanLabellingDocumentPageColumnHeader
    {

        private string propertyNameField;

        private string dataTypeField;

        /// <remarks/>
        public string PropertyName
        {
            get
            {
                return this.propertyNameField;
            }
            set
            {
                this.propertyNameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string DataType
        {
            get
            {
                return this.dataTypeField;
            }
            set
            {
                this.dataTypeField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EplanLabellingDocumentPageLine : IComparable
    {

        private EplanLabellingDocumentPageLineLabel labelField;

        private ushort source_idField;

        private string separatorField;
        public int CompareTo(EplanLabellingDocumentPageLine comparePart)
        {
            // A null value means that this object is greater.
            if (comparePart == null)
                return 1;
            else
            {
                if (this.Label.CompareTo(comparePart.Label) != 0)
                {
                    return this.Label.CompareTo(comparePart.Label);
                }
                else
                {
                    return 0;
                }
            }
        }

        public override int GetHashCode()
        {
            return Label.GetHashCode();
        }

        public bool Equals(EplanLabellingDocumentPageLine other)
        {
            if (other == null) return false;
            return (this.Label.Equals(other.Label));
        }

        public int CompareTo(object obj)
        {
            if (obj == null) return 1;

            EplanLabellingDocumentPageLine otherPageLine = obj as EplanLabellingDocumentPageLine;
            if (otherPageLine != null)
                return this.Label.CompareTo(otherPageLine.Label);
            else
                throw new ArgumentException("Object is not a PageLine");
        }


        /// <remarks/>
        public EplanLabellingDocumentPageLineLabel Label
        {
            get
            {
                return this.labelField;
            }
            set
            {
                this.labelField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ushort source_id
        {
            get
            {
                return this.source_idField;
            }
            set
            {
                this.source_idField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public string separator
        {
            get
            {
                return this.separatorField;
            }
            set
            {
                this.separatorField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EplanLabellingDocumentPageLineLabel
    {

        private EplanLabellingDocumentPageLineLabelProperty[] propertyField;

        private ushort source_idField;

        public int CompareTo(EplanLabellingDocumentPageLineLabel comparePart)
        {
            // A null value means that this object is greater.
            if (comparePart == null)
                return 1;
            else
            {
                return RecursiveSort(comparePart, 1);
            }
        }

        /// <summary>
        /// Recursive sorting to compare 6 properies
        /// </summary>
        /// <param name="comparePart"></param>
        /// <param name="v"></param>
        /// <returns></returns>
        private int RecursiveSort(EplanLabellingDocumentPageLineLabel comparePart, int v)
        {
            if (v == 7)
            {
                return 0;
            }
            if (this.Property[v].CompareTo(comparePart.Property[v]) != 0)
            {
                return this.Property[v].CompareTo(comparePart.Property[v]);
            }
            else
            {
                return RecursiveSort(comparePart, v + 1);
            }
        }

        public override int GetHashCode()
        {
            return Property[1].GetHashCode();
        }
        public bool Equals(EplanLabellingDocumentPageLineLabel other)
        {
            if (other == null) return false;
            return (this.Property[1].Equals(other.Property[1]));
        }


        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("Property")]
        public EplanLabellingDocumentPageLineLabelProperty[] Property
        {
            get
            {
                return this.propertyField;
            }
            set
            {
                this.propertyField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public ushort source_id
        {
            get
            {
                return this.source_idField;
            }
            set
            {
                this.source_idField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EplanLabellingDocumentPageLineLabelProperty
    {

        private string propertyNameField;

        private string propertyValueField;

        private byte formattingTypeField;

        private byte formattingLengthField;

        private byte formattingRAlignField;


        public int CompareTo(EplanLabellingDocumentPageLineLabelProperty comparePart)
        {
            // A null value means that this object is greater.
            if (comparePart == null)
                return 1;
            else
            {
                int numberPropertyValue;
                int numberComparePropertyValue;
                if (int.TryParse(this.PropertyValue, out numberPropertyValue) && int.TryParse(comparePart.PropertyValue, out numberComparePropertyValue) && numberPropertyValue.CompareTo(numberComparePropertyValue) != 0)
                {
                    return numberPropertyValue.CompareTo(numberComparePropertyValue);
                }                
                else if (this.PropertyValue.CompareTo(comparePart.PropertyValue) != 0)
                {
                    return this.PropertyValue.CompareTo(comparePart.PropertyValue);
                }
                else
                {
                    return 0;
                }
            }
        }
        public override int GetHashCode()
        {
            return PropertyValue.GetHashCode();
        }
        public bool Equals(EplanLabellingDocumentPageLineLabelProperty other)
        {
            if (other == null) return false;
            return (this.PropertyValue.Equals(other.PropertyValue));
        }

        /// <remarks/>
        public string PropertyName
        {
            get
            {
                return this.propertyNameField;
            }
            set
            {
                this.propertyNameField = value;
            }
        }

        /// <remarks/>
        public string PropertyValue
        {
            get
            {
                return this.propertyValueField;
            }
            set
            {
                this.propertyValueField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte FormattingType
        {
            get
            {
                return this.formattingTypeField;
            }
            set
            {
                this.formattingTypeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte FormattingLength
        {
            get
            {
                return this.formattingLengthField;
            }
            set
            {
                this.formattingLengthField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte FormattingRAlign
        {
            get
            {
                return this.formattingRAlignField;
            }
            set
            {
                this.formattingRAlignField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EplanLabellingDocumentPageFooter
    {

        private EplanLabellingDocumentPageFooterProperty propertyField;

        /// <remarks/>
        public EplanLabellingDocumentPageFooterProperty Property
        {
            get
            {
                return this.propertyField;
            }
            set
            {
                this.propertyField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class EplanLabellingDocumentPageFooterProperty
    {

        private string propertyNameField;

        private string propertyValueField;

        private byte formattingTypeField;

        private byte formattingLengthField;

        private byte formattingRAlignField;

        /// <remarks/>
        public string PropertyName
        {
            get
            {
                return this.propertyNameField;
            }
            set
            {
                this.propertyNameField = value;
            }
        }

        /// <remarks/>
        public string PropertyValue
        {
            get
            {
                return this.propertyValueField;
            }
            set
            {
                this.propertyValueField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte FormattingType
        {
            get
            {
                return this.formattingTypeField;
            }
            set
            {
                this.formattingTypeField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte FormattingLength
        {
            get
            {
                return this.formattingLengthField;
            }
            set
            {
                this.formattingLengthField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlAttributeAttribute()]
        public byte FormattingRAlign
        {
            get
            {
                return this.formattingRAlignField;
            }
            set
            {
                this.formattingRAlignField = value;
            }
        }
    }


}
