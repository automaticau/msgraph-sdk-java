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
 * The class for the Workbook Range Cell Parameter Set.
 */
public class WorkbookRangeCellParameterSet {
    /**
     * The row.
     * 
     */
    @SerializedName(value = "row", alternate = {"Row"})
    @Expose
	@Nullable
    public Integer row;

    /**
     * The column.
     * 
     */
    @SerializedName(value = "column", alternate = {"Column"})
    @Expose
	@Nullable
    public Integer column;


    /**
     * Instiaciates a new WorkbookRangeCellParameterSet
     */
    public WorkbookRangeCellParameterSet() {}
    /**
     * Instiaciates a new WorkbookRangeCellParameterSet
     * @param builder builder bearing the parameters to initialize from
     */
    protected WorkbookRangeCellParameterSet(@Nonnull final WorkbookRangeCellParameterSetBuilder builder) {
        this.row = builder.row;
        this.column = builder.column;
    }
    /**
     * Gets a new builder for the body
     * @return a new builder
     */
    @Nonnull
    public static WorkbookRangeCellParameterSetBuilder newBuilder() {
        return new WorkbookRangeCellParameterSetBuilder();
    }
    /**
     * Fluent builder for the WorkbookRangeCellParameterSet
     */
    public static final class WorkbookRangeCellParameterSetBuilder {
        /**
         * The row parameter value
         */
        @Nullable
        protected Integer row;
        /**
         * Sets the Row
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookRangeCellParameterSetBuilder withRow(@Nullable final Integer val) {
            this.row = val;
            return this;
        }
        /**
         * The column parameter value
         */
        @Nullable
        protected Integer column;
        /**
         * Sets the Column
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookRangeCellParameterSetBuilder withColumn(@Nullable final Integer val) {
            this.column = val;
            return this;
        }
        /**
         * Instanciates a new WorkbookRangeCellParameterSetBuilder
         */
        @Nullable
        protected WorkbookRangeCellParameterSetBuilder(){}
        /**
         * Buils the resulting body object to be passed to the request
         * @return the body object to pass to the request
         */
        @Nonnull
        public WorkbookRangeCellParameterSet build() {
            return new WorkbookRangeCellParameterSet(this);
        }
    }
    /**
     * Gets the functions options from the properties that have been set
     * @return a list of function options for the request
     */
    @Nonnull
    public java.util.List<com.microsoft.graph.options.FunctionOption> getFunctionOptions() {
        final ArrayList<com.microsoft.graph.options.FunctionOption> result = new ArrayList<>();
        if(this.row != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("row", row));
        }
        if(this.column != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("column", column));
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
