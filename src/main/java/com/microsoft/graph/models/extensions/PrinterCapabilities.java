// Template Source: BaseEntity.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models.extensions;
import com.microsoft.graph.serializer.ISerializer;
import com.microsoft.graph.serializer.IJsonBackedObject;
import com.microsoft.graph.serializer.AdditionalDataManager;
import java.util.EnumSet;
import com.microsoft.graph.models.generated.PrintColorMode;
import com.microsoft.graph.models.extensions.IntegerRange;
import com.microsoft.graph.models.generated.PrintDuplexMode;
import com.microsoft.graph.models.generated.PrinterFeedOrientation;
import com.microsoft.graph.models.generated.PrintFinishing;
import com.microsoft.graph.models.generated.PrintMultipageLayout;
import com.microsoft.graph.models.generated.PrintOrientation;
import com.microsoft.graph.models.generated.PrintQuality;
import com.microsoft.graph.models.generated.PrintScaling;


import com.google.gson.JsonObject;
import com.google.gson.annotations.SerializedName;
import com.google.gson.annotations.Expose;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Printer Capabilities.
 */
public class PrinterCapabilities implements IJsonBackedObject {

    @SerializedName("@odata.type")
    @Expose
    public String oDataType;

    private transient AdditionalDataManager additionalDataManager = new AdditionalDataManager(this);

    @Override
    public final AdditionalDataManager additionalDataManager() {
        return additionalDataManager;
    }

    /**
     * The Bottom Margins.
     * A list of supported bottom margins(in microns) for the printer.
     */
    @SerializedName(value = "bottomMargins", alternate = {"BottomMargins"})
    @Expose
    public java.util.List<Integer> bottomMargins;

    /**
     * The Collation.
     * True if the printer supports collating when printing muliple copies of a multi-page document; false otherwise.
     */
    @SerializedName(value = "collation", alternate = {"Collation"})
    @Expose
    public Boolean collation;

    /**
     * The Color Modes.
     * The color modes supported by the printer. Valid values are described in the following table.
     */
    @SerializedName(value = "colorModes", alternate = {"ColorModes"})
    @Expose
    public java.util.List<PrintColorMode> colorModes;

    /**
     * The Content Types.
     * A list of supported content (MIME) types that the printer supports. It is not guaranteed that the Universal Print service supports printing all of these MIME types.
     */
    @SerializedName(value = "contentTypes", alternate = {"ContentTypes"})
    @Expose
    public java.util.List<String> contentTypes;

    /**
     * The Copies Per Job.
     * The range of copies per job supported by the printer.
     */
    @SerializedName(value = "copiesPerJob", alternate = {"CopiesPerJob"})
    @Expose
    public IntegerRange copiesPerJob;

    /**
     * The Dpis.
     * The list of print resolutions in DPI that are supported by the printer.
     */
    @SerializedName(value = "dpis", alternate = {"Dpis"})
    @Expose
    public java.util.List<Integer> dpis;

    /**
     * The Duplex Modes.
     * The list of duplex modes that are supported by the printer. Valid values are described in the following table.
     */
    @SerializedName(value = "duplexModes", alternate = {"DuplexModes"})
    @Expose
    public java.util.List<PrintDuplexMode> duplexModes;

    /**
     * The Feed Orientations.
     * The list of feed orientations that are supported by the printer.
     */
    @SerializedName(value = "feedOrientations", alternate = {"FeedOrientations"})
    @Expose
    public java.util.List<PrinterFeedOrientation> feedOrientations;

    /**
     * The Finishings.
     * Finishing processes the printer supports for a printed document.
     */
    @SerializedName(value = "finishings", alternate = {"Finishings"})
    @Expose
    public java.util.List<PrintFinishing> finishings;

    /**
     * The Input Bins.
     * Supported input bins for the printer.
     */
    @SerializedName(value = "inputBins", alternate = {"InputBins"})
    @Expose
    public java.util.List<String> inputBins;

    /**
     * The Is Color Printing Supported.
     * True if color printing is supported by the printer; false otherwise. Read-only.
     */
    @SerializedName(value = "isColorPrintingSupported", alternate = {"IsColorPrintingSupported"})
    @Expose
    public Boolean isColorPrintingSupported;

    /**
     * The Is Page Range Supported.
     * True if the printer supports printing by page ranges; false otherwise.
     */
    @SerializedName(value = "isPageRangeSupported", alternate = {"IsPageRangeSupported"})
    @Expose
    public Boolean isPageRangeSupported;

    /**
     * The Left Margins.
     * A list of supported left margins(in microns) for the printer.
     */
    @SerializedName(value = "leftMargins", alternate = {"LeftMargins"})
    @Expose
    public java.util.List<Integer> leftMargins;

    /**
     * The Media Colors.
     * The media (i.e., paper) colors supported by the printer.
     */
    @SerializedName(value = "mediaColors", alternate = {"MediaColors"})
    @Expose
    public java.util.List<String> mediaColors;

    /**
     * The Media Sizes.
     * The media sizes supported by the printer. Supports standard size names for ISO and ANSI media sizes, along with any custom sizes supported by the associated printer.
     */
    @SerializedName(value = "mediaSizes", alternate = {"MediaSizes"})
    @Expose
    public java.util.List<String> mediaSizes;

    /**
     * The Media Types.
     * The media types supported by the printer. Valid values are described in the following table.
     */
    @SerializedName(value = "mediaTypes", alternate = {"MediaTypes"})
    @Expose
    public java.util.List<String> mediaTypes;

    /**
     * The Multipage Layouts.
     * The presentation directions supported by the printer. Supported values are described in the following table.
     */
    @SerializedName(value = "multipageLayouts", alternate = {"MultipageLayouts"})
    @Expose
    public java.util.List<PrintMultipageLayout> multipageLayouts;

    /**
     * The Orientations.
     * The print orientations supported by the printer. Valid values are described in the following table.
     */
    @SerializedName(value = "orientations", alternate = {"Orientations"})
    @Expose
    public java.util.List<PrintOrientation> orientations;

    /**
     * The Output Bins.
     * The printer's supported output bins (trays).
     */
    @SerializedName(value = "outputBins", alternate = {"OutputBins"})
    @Expose
    public java.util.List<String> outputBins;

    /**
     * The Pages Per Sheet.
     * Supported number of Input Pages to impose upon a single Impression.
     */
    @SerializedName(value = "pagesPerSheet", alternate = {"PagesPerSheet"})
    @Expose
    public java.util.List<Integer> pagesPerSheet;

    /**
     * The Qualities.
     * The print qualities supported by the printer.
     */
    @SerializedName(value = "qualities", alternate = {"Qualities"})
    @Expose
    public java.util.List<PrintQuality> qualities;

    /**
     * The Right Margins.
     * A list of supported right margins(in microns) for the printer.
     */
    @SerializedName(value = "rightMargins", alternate = {"RightMargins"})
    @Expose
    public java.util.List<Integer> rightMargins;

    /**
     * The Scalings.
     * Supported print scalings.
     */
    @SerializedName(value = "scalings", alternate = {"Scalings"})
    @Expose
    public java.util.List<PrintScaling> scalings;

    /**
     * The Supports Fit Pdf To Page.
     * True if the printer supports scaling PDF pages to match the print media size; false otherwise.
     */
    @SerializedName(value = "supportsFitPdfToPage", alternate = {"SupportsFitPdfToPage"})
    @Expose
    public Boolean supportsFitPdfToPage;

    /**
     * The Top Margins.
     * A list of supported top margins(in microns) for the printer.
     */
    @SerializedName(value = "topMargins", alternate = {"TopMargins"})
    @Expose
    public java.util.List<Integer> topMargins;


    /**
     * The raw representation of this class
     */
    private JsonObject rawObject;

    /**
     * The serializer
     */
    private ISerializer serializer;

    /**
     * Gets the raw representation of this class
     *
     * @return the raw representation of this class
     */
    public JsonObject getRawObject() {
        return rawObject;
    }

    /**
     * Gets serializer
     *
     * @return the serializer
     */
    protected ISerializer getSerializer() {
        return serializer;
    }

    /**
     * Sets the raw JSON object
     *
     * @param serializer the serializer
     * @param json the JSON object to set this object to
     */
    public void setRawObject(final ISerializer serializer, final JsonObject json) {
        this.serializer = serializer;
        rawObject = json;

    }
}