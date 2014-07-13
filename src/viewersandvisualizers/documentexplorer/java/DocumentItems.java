/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package viewersandvisualizers.documentexplorer.java;

import com.aspose.words.*;

import java.text.MessageFormat;

public class DocumentItems {
    // Classes inherited from the Item class provide specialized representation of particular
    // document node by overriding virtual methods and properties of the base class.

    // The NODE_TYPE_STRING field is used to build the map between the classes and the names of the corresponding node types.
    public class DocumentItem extends Item {

        public DocumentItem(Node node) {
            super(node);
        }

        public boolean isRemovable() {
            return false;
        }
        public static final String NODE_TYPE_STRING = "DOCUMENT";
    }

    public class SectionItem extends Item {

        public SectionItem(Node node) {
            super(node);
        }

        public boolean isRemovable() {
            if (this.getNode() == ((Document) this.getNode().getDocument()).getLastSection()) {
                return false;
            } else {
                return true;
            }
        }
        public static final String NODE_TYPE_STRING = "SECTION";
    }

    public class HeaderFooterItem extends Item {

        public HeaderFooterItem(Node node) {
            super(node);
        }

        protected String getIconName() throws Exception {
            if (((HeaderFooter) this.getNode()).isHeader()) {
                return "Header";
            } else {
                return "Footer";
            }
        }

        public String getName() throws Exception {
            return MessageFormat.format("{0} - {1}", super.getName(), getHeaderFooterTypeAsString((HeaderFooter) this.getNode()));
        }
        public static final String NODE_TYPE_STRING = "HEADER_FOOTER";
    }

    public class BodyItem extends Item {

        public BodyItem(Node node) {
            super(node);
        }

        public boolean isRemovable() {
            return false;
        }
        public static final String NODE_TYPE_STRING = "BODY";
    }

    public class TableItem extends Item {

        public TableItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "TABLE";
    }

    public class RowItem extends Item {

        public RowItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "ROW";
    }

    public class CellItem extends Item {

        public CellItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "CELL";
    }

    public class ParagraphItem extends Item {

        public ParagraphItem(Node node) {
            super(node);
        }

        public boolean isRemovable() {
            Paragraph para = (Paragraph) this.getNode();
            if (para.isEndOfSection()) {
                return false;
            } else {
                return true;
            }
        }
        public static final String NODE_TYPE_STRING = "PARAGRAPH";
    }

    public class RunItem extends Item {

        public RunItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "RUN";
    }

    public class FieldStartItem extends Item {

        public FieldStartItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "FIELD_START";
    }

    public class FieldSeparatorItem extends Item {

        public FieldSeparatorItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "FIELD_SEPARATOR";
    }

    public class FieldEndItem extends Item {

        public FieldEndItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "FIELD_END";
    }

    public class BookmarkStartItem extends Item {

        public BookmarkStartItem(Node node) {
            super(node);
        }

        public String getName() throws Exception {
            return MessageFormat.format("{0} - \"{1}\"", super.getName(), ((BookmarkStart) this.getNode()).getName());
        }
        public static final String NODE_TYPE_STRING = "BOOKMARK_START";
    }

    public class BookmarkEndItem extends Item {

        public BookmarkEndItem(Node node) {
            super(node);
        }

        public String getName() throws Exception {
            return MessageFormat.format("{0} - \"{1}\"", super.getName(), ((BookmarkEnd) this.getNode()).getName());
        }
        public static final String NODE_TYPE_STRING = "BOOKMARK_END";
    }

    public class CustomXmlMarkupItem extends Item {

        public CustomXmlMarkupItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "CUSTOM_XML_MARKUP";
    }

    public class StructuredDocumentTagItem extends Item {

        public StructuredDocumentTagItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "STRUCTURED_DOCUMENT_TAG";
    }

    public class CommentItem extends Item {

        public String getName() throws Exception {
            return String.format("%s - (Id = %s)", super.getName(), ((Comment) getNode()).getId());
        }

        public CommentItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "COMMENT";
    }

    public class CommentRangeStartItem extends Item {

        public String getName() throws Exception {
            return String.format("%s - (Id = %s)", super.getName(), ((CommentRangeStart) getNode()).getId());
        }

        public CommentRangeStartItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "COMMENT_RANGE_START";
    }

    public class CommentRangeEndItem extends Item {

        public String getName() throws Exception {
            return String.format("%s - (Id = %s)", super.getName(), ((CommentRangeEnd) getNode()).getId());
        }

        public CommentRangeEndItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "COMMENT_RANGE_END";
    }

    public class DrawingMLItem extends Item {

        public DrawingMLItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "DRAWING_ML";
    }

    public class OfficeMathItem extends Item {

        public OfficeMathItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "OFFICE_MATH";
    }

    public class SmartTagItem extends Item {

        public SmartTagItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "SMART_TAG";
    }

    public class GroupShapeItem extends Item {

        public GroupShapeItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "GROUP_SHAPE";
    }

    public class FootnoteItem extends Item {

        public FootnoteItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "FOOTNOTE";
    }

    public class ShapeItem extends Item {

        public ShapeItem(Node node) {
            super(node);
        }

        public String getName() throws Exception {
            Shape shape = (Shape) getNode();

            switch (shape.getShapeType()) {
                case ShapeType.OLE_OBJECT:
                    return shape.getOleFormat().getProgId();
                case ShapeType.OLE_CONTROL:
                    return shape.getOleFormat().getProgId();
                default:
                    return super.getIconName();
            }
        }

        protected String getIconName() throws Exception {
            Shape shape = (Shape) getNode();

            switch (shape.getShapeType()) {
                case ShapeType.OLE_OBJECT:
                    return "OleObject";
                case ShapeType.OLE_CONTROL:
                    return "OleControl";
                default:
                    if (shape.isInline()) {
                        return "InlineShape";
                    } else {
                        return super.getIconName();
                    }
            }
        }
        public static final String NODE_TYPE_STRING = "SHAPE";
    }

    public class FormFieldItem extends Item {

        public FormFieldItem(Node node) {
            super(node);
        }

        public String getName() throws Exception {
            return MessageFormat.format("{0} - \"{1}\"", super.getName(), ((FormField) this.getNode()).getName());
        }

        protected String getIconName() throws Exception {
            switch (((FormField) this.getNode()).getType()) {
                case FieldType.FIELD_FORM_CHECK_BOX:
                    return "FormCheckBox";
                case FieldType.FIELD_FORM_DROP_DOWN:
                    return "FormDropDown";
                case FieldType.FIELD_FORM_TEXT_INPUT:
                    return "FormTextInput";
                default:
                    return super.getIconName();
            }
        }
        public static final String NODE_TYPE_STRING = "FORM_FIELD";
    }

    public class SpecialCharItem extends Item {

        public SpecialCharItem(Node node) {
            super(node);
        }
        public static final String NODE_TYPE_STRING = "SPECIAL_CHAR";
    }
}
