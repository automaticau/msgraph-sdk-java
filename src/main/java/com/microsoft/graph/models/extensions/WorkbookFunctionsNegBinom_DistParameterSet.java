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
 * The class for the Workbook Functions Neg Binom_Dist Parameter Set.
 */
public class WorkbookFunctionsNegBinom_DistParameterSet {
    /**
     * The number F.
     * 
     */
    @SerializedName(value = "numberF", alternate = {"NumberF"})
    @Expose
	@Nullable
    public com.google.gson.JsonElement numberF;

    /**
     * The number S.
     * 
     */
    @SerializedName(value = "numberS", alternate = {"NumberS"})
    @Expose
	@Nullable
    public com.google.gson.JsonElement numberS;

    /**
     * The probability S.
     * 
     */
    @SerializedName(value = "probabilityS", alternate = {"ProbabilityS"})
    @Expose
	@Nullable
    public com.google.gson.JsonElement probabilityS;

    /**
     * The cumulative.
     * 
     */
    @SerializedName(value = "cumulative", alternate = {"Cumulative"})
    @Expose
	@Nullable
    public com.google.gson.JsonElement cumulative;


    /**
     * Instiaciates a new WorkbookFunctionsNegBinom_DistParameterSet
     */
    public WorkbookFunctionsNegBinom_DistParameterSet() {}
    /**
     * Instiaciates a new WorkbookFunctionsNegBinom_DistParameterSet
     * @param builder builder bearing the parameters to initialize from
     */
    protected WorkbookFunctionsNegBinom_DistParameterSet(@Nonnull final WorkbookFunctionsNegBinom_DistParameterSetBuilder builder) {
        this.numberF = builder.numberF;
        this.numberS = builder.numberS;
        this.probabilityS = builder.probabilityS;
        this.cumulative = builder.cumulative;
    }
    /**
     * Gets a new builder for the body
     * @return a new builder
     */
    @Nonnull
    public static WorkbookFunctionsNegBinom_DistParameterSetBuilder newBuilder() {
        return new WorkbookFunctionsNegBinom_DistParameterSetBuilder();
    }
    /**
     * Fluent builder for the WorkbookFunctionsNegBinom_DistParameterSet
     */
    public static final class WorkbookFunctionsNegBinom_DistParameterSetBuilder {
        /**
         * The numberF parameter value
         */
        @Nullable
        protected com.google.gson.JsonElement numberF;
        /**
         * Sets the NumberF
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookFunctionsNegBinom_DistParameterSetBuilder withNumberF(@Nullable final com.google.gson.JsonElement val) {
            this.numberF = val;
            return this;
        }
        /**
         * The numberS parameter value
         */
        @Nullable
        protected com.google.gson.JsonElement numberS;
        /**
         * Sets the NumberS
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookFunctionsNegBinom_DistParameterSetBuilder withNumberS(@Nullable final com.google.gson.JsonElement val) {
            this.numberS = val;
            return this;
        }
        /**
         * The probabilityS parameter value
         */
        @Nullable
        protected com.google.gson.JsonElement probabilityS;
        /**
         * Sets the ProbabilityS
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookFunctionsNegBinom_DistParameterSetBuilder withProbabilityS(@Nullable final com.google.gson.JsonElement val) {
            this.probabilityS = val;
            return this;
        }
        /**
         * The cumulative parameter value
         */
        @Nullable
        protected com.google.gson.JsonElement cumulative;
        /**
         * Sets the Cumulative
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public WorkbookFunctionsNegBinom_DistParameterSetBuilder withCumulative(@Nullable final com.google.gson.JsonElement val) {
            this.cumulative = val;
            return this;
        }
        /**
         * Instanciates a new WorkbookFunctionsNegBinom_DistParameterSetBuilder
         */
        @Nullable
        protected WorkbookFunctionsNegBinom_DistParameterSetBuilder(){}
        /**
         * Buils the resulting body object to be passed to the request
         * @return the body object to pass to the request
         */
        @Nonnull
        public WorkbookFunctionsNegBinom_DistParameterSet build() {
            return new WorkbookFunctionsNegBinom_DistParameterSet(this);
        }
    }
    /**
     * Gets the functions options from the properties that have been set
     * @return a list of function options for the request
     */
    @Nonnull
    public java.util.List<com.microsoft.graph.options.FunctionOption> getFunctionOptions() {
        final ArrayList<com.microsoft.graph.options.FunctionOption> result = new ArrayList<>();
        if(this.numberF != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("numberF", numberF));
        }
        if(this.numberS != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("numberS", numberS));
        }
        if(this.probabilityS != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("probabilityS", probabilityS));
        }
        if(this.cumulative != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("cumulative", cumulative));
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
