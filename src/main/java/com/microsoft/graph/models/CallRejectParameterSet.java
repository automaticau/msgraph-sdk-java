// Template Source: BaseMethodParameterSet.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models;

import com.microsoft.graph.models.RejectReason;
import com.google.gson.annotations.Expose;
import com.google.gson.annotations.SerializedName;
import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import com.google.gson.JsonObject;
import java.util.EnumSet;
import java.util.ArrayList;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Call Reject Parameter Set.
 */
public class CallRejectParameterSet {
    /**
     * The reason.
     * 
     */
    @SerializedName(value = "reason", alternate = {"Reason"})
    @Expose
	@Nullable
    public RejectReason reason;

    /**
     * The callback Uri.
     * 
     */
    @SerializedName(value = "callbackUri", alternate = {"CallbackUri"})
    @Expose
	@Nullable
    public String callbackUri;


    /**
     * Instiaciates a new CallRejectParameterSet
     */
    public CallRejectParameterSet() {}
    /**
     * Instiaciates a new CallRejectParameterSet
     * @param builder builder bearing the parameters to initialize from
     */
    protected CallRejectParameterSet(@Nonnull final CallRejectParameterSetBuilder builder) {
        this.reason = builder.reason;
        this.callbackUri = builder.callbackUri;
    }
    /**
     * Gets a new builder for the body
     * @return a new builder
     */
    @Nonnull
    public static CallRejectParameterSetBuilder newBuilder() {
        return new CallRejectParameterSetBuilder();
    }
    /**
     * Fluent builder for the CallRejectParameterSet
     */
    public static final class CallRejectParameterSetBuilder {
        /**
         * The reason parameter value
         */
        @Nullable
        protected RejectReason reason;
        /**
         * Sets the Reason
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public CallRejectParameterSetBuilder withReason(@Nullable final RejectReason val) {
            this.reason = val;
            return this;
        }
        /**
         * The callbackUri parameter value
         */
        @Nullable
        protected String callbackUri;
        /**
         * Sets the CallbackUri
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public CallRejectParameterSetBuilder withCallbackUri(@Nullable final String val) {
            this.callbackUri = val;
            return this;
        }
        /**
         * Instanciates a new CallRejectParameterSetBuilder
         */
        @Nullable
        protected CallRejectParameterSetBuilder(){}
        /**
         * Buils the resulting body object to be passed to the request
         * @return the body object to pass to the request
         */
        @Nonnull
        public CallRejectParameterSet build() {
            return new CallRejectParameterSet(this);
        }
    }
    /**
     * Gets the functions options from the properties that have been set
     * @return a list of function options for the request
     */
    @Nonnull
    public java.util.List<com.microsoft.graph.options.FunctionOption> getFunctionOptions() {
        final ArrayList<com.microsoft.graph.options.FunctionOption> result = new ArrayList<>();
        if(this.reason != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("reason", reason));
        }
        if(this.callbackUri != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("callbackUri", callbackUri));
        }
        return result;
    }
}
