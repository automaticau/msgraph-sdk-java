// Template Source: BaseMethodParameterSet.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models.extensions;

import com.microsoft.graph.models.extensions.WorkbookFunctionResult;
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
 * The class for the Workbook Functions Convert Parameter Set.
 */
public class WorkbookFunctionsConvertParameterSet {
    /**
     * The number.
     * 
     */
    @SerializedName(value = "number", alternate = {"Number"})
    @Expose
	@Nullable
    public com.google.gson.JsonElement number;

    /**
     * The from Unit.
     * 
     */
    @SerializedName(value = "fromUnit", alternate = {"FromUnit"})
    @Expose
	@Nullable
    public com.google.gson.JsonElement fromUnit;

    /**
     * The to Unit.
     * 
     */
    @SerializedName(value = "toUnit", alternate = {"ToUnit"})
    @Expose
	@Nullable
    public com.google.gson.JsonElement toUnit;


    /**
     * Instiaciates a new WorkbookFunctionsConvertParameterSet
     */
    public WorkbookFunctionsConvertParameterSet() {}
    /**
     * Instiaciates a new WorkbookFunctionsConvertParameterSet
     * @param builder builder bearing the parameters to initialize from
     */
    protected WorkbookFunctionsConvertParameterSet(@Nonnull final WorkbookFunctionsConvertParameterSetBuilder builder) {
        this.number = builder.number;
        this.fromUnit = builder.fromUnit;
        this.toUnit = builder.toUnit;
    }
    /**
     * Gets a new builder for the body
     * @return a new builder
     */
    @Nonnull
    public static WorkbookFunctionsConvertParameterSetBuilder newBuilder() {
        return new WorkbookFunctionsConvertParameterSetBuilder();
    }
    /**
     * Fluent builder for the WorkbookFunctionsConvertParameterSet
     */
    public static final class WorkbookFunctionsConvertParameterSetBuilder {
        /**
         * The number parameter value
         */
        @Nullable
        protected com.google.gson.JsonElement number;
        /**
         * Sets the Number
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookFunctionsConvertParameterSetBuilder withNumber(@Nullable final com.google.gson.JsonElement val) {
            this.number = val;
            return this;
        }
        /**
         * The fromUnit parameter value
         */
        @Nullable
        protected com.google.gson.JsonElement fromUnit;
        /**
         * Sets the FromUnit
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookFunctionsConvertParameterSetBuilder withFromUnit(@Nullable final com.google.gson.JsonElement val) {
            this.fromUnit = val;
            return this;
        }
        /**
         * The toUnit parameter value
         */
        @Nullable
        protected com.google.gson.JsonElement toUnit;
        /**
         * Sets the ToUnit
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookFunctionsConvertParameterSetBuilder withToUnit(@Nullable final com.google.gson.JsonElement val) {
            this.toUnit = val;
            return this;
        }
        /**
         * Instanciates a new WorkbookFunctionsConvertParameterSetBuilder
         */
        @Nullable
        protected WorkbookFunctionsConvertParameterSetBuilder(){}
        /**
         * Buils the resulting body object to be passed to the request
         * @return the body object to pass to the request
         */
        @Nonnull
        public WorkbookFunctionsConvertParameterSet build() {
            return new WorkbookFunctionsConvertParameterSet(this);
        }
    }
    /**
     * Gets the functions options from the properties that have been set
     * @return a list of function options for the request
     */
    @Nonnull
    public java.util.List<com.microsoft.graph.options.FunctionOption> getFunctionOptions() {
        final ArrayList<com.microsoft.graph.options.FunctionOption> result = new ArrayList<>();
        if(this.number != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("number", number));
        }
        if(this.fromUnit != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("fromUnit", fromUnit));
        }
        if(this.toUnit != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("toUnit", toUnit));
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
