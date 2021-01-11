// Template Source: BaseMethodParameterSet.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models.extensions;

import com.microsoft.graph.models.extensions.Report;
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
 * The class for the Report Root Get One Drive Usage Storage Parameter Set.
 */
public class ReportRootGetOneDriveUsageStorageParameterSet {
    /**
     * The period.
     * 
     */
    @SerializedName(value = "period", alternate = {"Period"})
    @Expose
	@Nullable
    public String period;


    /**
     * Instiaciates a new ReportRootGetOneDriveUsageStorageParameterSet
     */
    public ReportRootGetOneDriveUsageStorageParameterSet() {}
    /**
     * Instiaciates a new ReportRootGetOneDriveUsageStorageParameterSet
     * @param builder builder bearing the parameters to initialize from
     */
    protected ReportRootGetOneDriveUsageStorageParameterSet(@Nonnull final ReportRootGetOneDriveUsageStorageParameterSetBuilder builder) {
        this.period = builder.period;
    }
    /**
     * Gets a new builder for the body
     * @return a new builder
     */
    @Nonnull
    public static ReportRootGetOneDriveUsageStorageParameterSetBuilder newBuilder() {
        return new ReportRootGetOneDriveUsageStorageParameterSetBuilder();
    }
    /**
     * Fluent builder for the ReportRootGetOneDriveUsageStorageParameterSet
     */
    public static final class ReportRootGetOneDriveUsageStorageParameterSetBuilder {
        /**
         * The period parameter value
         */
        @Nullable
        protected String period;
        /**
         * Sets the Period
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public ReportRootGetOneDriveUsageStorageParameterSetBuilder withPeriod(@Nullable final String val) {
            this.period = val;
            return this;
        }
        /**
         * Instanciates a new ReportRootGetOneDriveUsageStorageParameterSetBuilder
         */
        @Nullable
        protected ReportRootGetOneDriveUsageStorageParameterSetBuilder(){}
        /**
         * Buils the resulting body object to be passed to the request
         * @return the body object to pass to the request
         */
        @Nonnull
        public ReportRootGetOneDriveUsageStorageParameterSet build() {
            return new ReportRootGetOneDriveUsageStorageParameterSet(this);
        }
    }
    /**
     * Gets the functions options from the properties that have been set
     * @return a list of function options for the request
     */
    @Nonnull
    public java.util.List<com.microsoft.graph.options.FunctionOption> getFunctionOptions() {
        final ArrayList<com.microsoft.graph.options.FunctionOption> result = new ArrayList<>();
        if(this.period != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("period", period));
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
