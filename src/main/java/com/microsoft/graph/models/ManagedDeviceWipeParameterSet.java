// Template Source: BaseMethodParameterSet.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models;


import com.google.gson.annotations.Expose;
import com.google.gson.annotations.SerializedName;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import com.google.gson.JsonObject;
import java.util.EnumSet;
import java.util.ArrayList;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Managed Device Wipe Parameter Set.
 */
public class ManagedDeviceWipeParameterSet {
    /**
     * The keep Enrollment Data.
     * 
     */
    @SerializedName(value = "keepEnrollmentData", alternate = {"KeepEnrollmentData"})
    @Expose
	@Nullable
    public Boolean keepEnrollmentData;

    /**
     * The keep User Data.
     * 
     */
    @SerializedName(value = "keepUserData", alternate = {"KeepUserData"})
    @Expose
	@Nullable
    public Boolean keepUserData;

    /**
     * The mac Os Unlock Code.
     * 
     */
    @SerializedName(value = "macOsUnlockCode", alternate = {"MacOsUnlockCode"})
    @Expose
	@Nullable
    public String macOsUnlockCode;

    /**
     * The persist Esim Data Plan.
     * 
     */
    @SerializedName(value = "persistEsimDataPlan", alternate = {"PersistEsimDataPlan"})
    @Expose
	@Nullable
    public Boolean persistEsimDataPlan;


    /**
     * Instiaciates a new ManagedDeviceWipeParameterSet
     */
    public ManagedDeviceWipeParameterSet() {}
    /**
     * Instiaciates a new ManagedDeviceWipeParameterSet
     * @param builder builder bearing the parameters to initialize from
     */
    protected ManagedDeviceWipeParameterSet(@Nonnull final ManagedDeviceWipeParameterSetBuilder builder) {
        this.keepEnrollmentData = builder.keepEnrollmentData;
        this.keepUserData = builder.keepUserData;
        this.macOsUnlockCode = builder.macOsUnlockCode;
        this.persistEsimDataPlan = builder.persistEsimDataPlan;
    }
    /**
     * Gets a new builder for the body
     * @return a new builder
     */
    @Nonnull
    public static ManagedDeviceWipeParameterSetBuilder newBuilder() {
        return new ManagedDeviceWipeParameterSetBuilder();
    }
    /**
     * Fluent builder for the ManagedDeviceWipeParameterSet
     */
    public static final class ManagedDeviceWipeParameterSetBuilder {
        /**
         * The keepEnrollmentData parameter value
         */
        @Nullable
        protected Boolean keepEnrollmentData;
        /**
         * Sets the KeepEnrollmentData
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public ManagedDeviceWipeParameterSetBuilder withKeepEnrollmentData(@Nullable final Boolean val) {
            this.keepEnrollmentData = val;
            return this;
        }
        /**
         * The keepUserData parameter value
         */
        @Nullable
        protected Boolean keepUserData;
        /**
         * Sets the KeepUserData
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public ManagedDeviceWipeParameterSetBuilder withKeepUserData(@Nullable final Boolean val) {
            this.keepUserData = val;
            return this;
        }
        /**
         * The macOsUnlockCode parameter value
         */
        @Nullable
        protected String macOsUnlockCode;
        /**
         * Sets the MacOsUnlockCode
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public ManagedDeviceWipeParameterSetBuilder withMacOsUnlockCode(@Nullable final String val) {
            this.macOsUnlockCode = val;
            return this;
        }
        /**
         * The persistEsimDataPlan parameter value
         */
        @Nullable
        protected Boolean persistEsimDataPlan;
        /**
         * Sets the PersistEsimDataPlan
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public ManagedDeviceWipeParameterSetBuilder withPersistEsimDataPlan(@Nullable final Boolean val) {
            this.persistEsimDataPlan = val;
            return this;
        }
        /**
         * Instanciates a new ManagedDeviceWipeParameterSetBuilder
         */
        @Nullable
        protected ManagedDeviceWipeParameterSetBuilder(){}
        /**
         * Buils the resulting body object to be passed to the request
         * @return the body object to pass to the request
         */
        @Nonnull
        public ManagedDeviceWipeParameterSet build() {
            return new ManagedDeviceWipeParameterSet(this);
        }
    }
    /**
     * Gets the functions options from the properties that have been set
     * @return a list of function options for the request
     */
    @Nonnull
    public java.util.List<com.microsoft.graph.options.FunctionOption> getFunctionOptions() {
        final ArrayList<com.microsoft.graph.options.FunctionOption> result = new ArrayList<>();
        if(this.keepEnrollmentData != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("keepEnrollmentData", keepEnrollmentData));
        }
        if(this.keepUserData != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("keepUserData", keepUserData));
        }
        if(this.macOsUnlockCode != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("macOsUnlockCode", macOsUnlockCode));
        }
        if(this.persistEsimDataPlan != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("persistEsimDataPlan", persistEsimDataPlan));
        }
        return result;
    }
}
