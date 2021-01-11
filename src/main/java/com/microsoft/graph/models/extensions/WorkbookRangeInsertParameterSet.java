// Template Source: BaseMethodParameterSet.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models.extensions;

import com.microsoft.graph.models.extensions.WorkbookRange;
import com.google.gson.annotations.Expose;
import com.google.gson.annotations.SerializedName;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import com.google.gson.JsonObject;
import com.microsoft.graph.serializer.ISerializer;
import java.util.EnumSet;
import java.util.ArrayList;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Workbook Range Insert Parameter Set.
 */
public class WorkbookRangeInsertParameterSet {
    /**
     * The shift.
     * 
     */
    @SerializedName(value = "shift", alternate = {"Shift"})
    @Expose
	@Nullable
    public String shift;


    /**
     * Instiaciates a new WorkbookRangeInsertParameterSet
     */
    public WorkbookRangeInsertParameterSet() {}
    /**
     * Instiaciates a new WorkbookRangeInsertParameterSet
     * @param builder builder bearing the parameters to initialize from
     */
    protected WorkbookRangeInsertParameterSet(@Nonnull final WorkbookRangeInsertParameterSetBuilder builder) {
        this.shift = builder.shift;
    }
    /**
     * Gets a new builder for the body
     * @return a new builder
     */
    @Nonnull
    public static WorkbookRangeInsertParameterSetBuilder newBuilder() {
        return new WorkbookRangeInsertParameterSetBuilder();
    }
    /**
     * Fluent builder for the WorkbookRangeInsertParameterSet
     */
    public static final class WorkbookRangeInsertParameterSetBuilder {
        /**
         * The shift parameter value
         */
        @Nullable
        protected String shift;
        /**
         * Sets the Shift
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookRangeInsertParameterSetBuilder withShift(@Nullable final String val) {
            this.shift = val;
            return this;
        }
        /**
         * Instanciates a new WorkbookRangeInsertParameterSetBuilder
         */
        @Nullable
        protected WorkbookRangeInsertParameterSetBuilder(){}
        /**
         * Buils the resulting body object to be passed to the request
         * @return the body object to pass to the request
         */
        @Nonnull
        public WorkbookRangeInsertParameterSet build() {
            return new WorkbookRangeInsertParameterSet(this);
        }
    }
    /**
     * Gets the functions options from the properties that have been set
     * @return a list of function options for the request
     */
    @Nonnull
    public java.util.List<com.microsoft.graph.options.FunctionOption> getFunctionOptions() {
        final ArrayList<com.microsoft.graph.options.FunctionOption> result = new ArrayList<>();
        if(this.shift != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("shift", shift));
        }
        return result;
    }
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
    @Nullable
    public JsonObject getRawObject() {
        return rawObject;
    }

    /**
     * Gets serializer
     *
     * @return the serializer
     */
    @Nullable
    public ISerializer getSerializer() {
        return serializer;
    }

    /**
     * Sets the raw JSON object
     *
     * @param serializer the serializer
     * @param json the JSON object to set this object to
     */
    public void setRawObject(@Nonnull final ISerializer serializer, @Nonnull final JsonObject json) {
        this.serializer = serializer;
        rawObject = json;

    }
}
