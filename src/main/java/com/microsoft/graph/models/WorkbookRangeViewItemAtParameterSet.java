// Template Source: BaseMethodParameterSet.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models;

import com.microsoft.graph.models.WorkbookRangeView;
import com.google.gson.annotations.Expose;
import com.google.gson.annotations.SerializedName;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import com.google.gson.JsonObject;
import java.util.EnumSet;
import java.util.ArrayList;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Workbook Range View Item At Parameter Set.
 */
public class WorkbookRangeViewItemAtParameterSet {
    /**
     * The index.
     * 
     */
    @SerializedName(value = "index", alternate = {"Index"})
    @Expose
	@Nullable
    public Integer index;


    /**
     * Instiaciates a new WorkbookRangeViewItemAtParameterSet
     */
    public WorkbookRangeViewItemAtParameterSet() {}
    /**
     * Instiaciates a new WorkbookRangeViewItemAtParameterSet
     * @param builder builder bearing the parameters to initialize from
     */
    protected WorkbookRangeViewItemAtParameterSet(@Nonnull final WorkbookRangeViewItemAtParameterSetBuilder builder) {
        this.index = builder.index;
    }
    /**
     * Gets a new builder for the body
     * @return a new builder
     */
    @Nonnull
    public static WorkbookRangeViewItemAtParameterSetBuilder newBuilder() {
        return new WorkbookRangeViewItemAtParameterSetBuilder();
    }
    /**
     * Fluent builder for the WorkbookRangeViewItemAtParameterSet
     */
    public static final class WorkbookRangeViewItemAtParameterSetBuilder {
        /**
         * The index parameter value
         */
        @Nullable
        protected Integer index;
        /**
         * Sets the Index
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookRangeViewItemAtParameterSetBuilder withIndex(@Nullable final Integer val) {
            this.index = val;
            return this;
        }
        /**
         * Instanciates a new WorkbookRangeViewItemAtParameterSetBuilder
         */
        @Nullable
        protected WorkbookRangeViewItemAtParameterSetBuilder(){}
        /**
         * Buils the resulting body object to be passed to the request
         * @return the body object to pass to the request
         */
        @Nonnull
        public WorkbookRangeViewItemAtParameterSet build() {
            return new WorkbookRangeViewItemAtParameterSet(this);
        }
    }
    /**
     * Gets the functions options from the properties that have been set
     * @return a list of function options for the request
     */
    @Nonnull
    public java.util.List<com.microsoft.graph.options.FunctionOption> getFunctionOptions() {
        final ArrayList<com.microsoft.graph.options.FunctionOption> result = new ArrayList<>();
        if(this.index != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("index", index));
        }
        return result;
    }
}
