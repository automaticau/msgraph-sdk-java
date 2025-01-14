// Template Source: BaseMethodParameterSet.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models;

import com.microsoft.graph.models.WorkbookRange;
import com.google.gson.annotations.Expose;
import com.google.gson.annotations.SerializedName;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import com.google.gson.JsonObject;
import java.util.EnumSet;
import java.util.ArrayList;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Workbook Worksheet Range Parameter Set.
 */
public class WorkbookWorksheetRangeParameterSet {
    /**
     * The address.
     * 
     */
    @SerializedName(value = "address", alternate = {"Address"})
    @Expose
	@Nullable
    public String address;


    /**
     * Instiaciates a new WorkbookWorksheetRangeParameterSet
     */
    public WorkbookWorksheetRangeParameterSet() {}
    /**
     * Instiaciates a new WorkbookWorksheetRangeParameterSet
     * @param builder builder bearing the parameters to initialize from
     */
    protected WorkbookWorksheetRangeParameterSet(@Nonnull final WorkbookWorksheetRangeParameterSetBuilder builder) {
        this.address = builder.address;
    }
    /**
     * Gets a new builder for the body
     * @return a new builder
     */
    @Nonnull
    public static WorkbookWorksheetRangeParameterSetBuilder newBuilder() {
        return new WorkbookWorksheetRangeParameterSetBuilder();
    }
    /**
     * Fluent builder for the WorkbookWorksheetRangeParameterSet
     */
    public static final class WorkbookWorksheetRangeParameterSetBuilder {
        /**
         * The address parameter value
         */
        @Nullable
        protected String address;
        /**
         * Sets the Address
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookWorksheetRangeParameterSetBuilder withAddress(@Nullable final String val) {
            this.address = val;
            return this;
        }
        /**
         * Instanciates a new WorkbookWorksheetRangeParameterSetBuilder
         */
        @Nullable
        protected WorkbookWorksheetRangeParameterSetBuilder(){}
        /**
         * Buils the resulting body object to be passed to the request
         * @return the body object to pass to the request
         */
        @Nonnull
        public WorkbookWorksheetRangeParameterSet build() {
            return new WorkbookWorksheetRangeParameterSet(this);
        }
    }
    /**
     * Gets the functions options from the properties that have been set
     * @return a list of function options for the request
     */
    @Nonnull
    public java.util.List<com.microsoft.graph.options.FunctionOption> getFunctionOptions() {
        final ArrayList<com.microsoft.graph.options.FunctionOption> result = new ArrayList<>();
        if(this.address != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("address", address));
        }
        return result;
    }
}
