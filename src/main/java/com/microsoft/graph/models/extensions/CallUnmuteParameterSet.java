// Template Source: BaseMethodParameterSet.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models.extensions;

import com.microsoft.graph.models.extensions.UnmuteParticipantOperation;
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
 * The class for the Call Unmute Parameter Set.
 */
public class CallUnmuteParameterSet {
    /**
     * The client Context.
     * 
     */
    @SerializedName(value = "clientContext", alternate = {"ClientContext"})
    @Expose
	@Nullable
    public String clientContext;


    /**
     * Instiaciates a new CallUnmuteParameterSet
     */
    public CallUnmuteParameterSet() {}
    /**
     * Instiaciates a new CallUnmuteParameterSet
     * @param builder builder bearing the parameters to initialize from
     */
    protected CallUnmuteParameterSet(@Nonnull final CallUnmuteParameterSetBuilder builder) {
        this.clientContext = builder.clientContext;
    }
    /**
     * Gets a new builder for the body
     * @return a new builder
     */
    @Nonnull
    public static CallUnmuteParameterSetBuilder newBuilder() {
        return new CallUnmuteParameterSetBuilder();
    }
    /**
     * Fluent builder for the CallUnmuteParameterSet
     */
    public static final class CallUnmuteParameterSetBuilder {
        /**
         * The clientContext parameter value
         */
        @Nullable
        protected String clientContext;
        /**
         * Sets the ClientContext
         * @param val the value to set it to
         * @return the current builder object
         */
        @Nonnull
        public CallUnmuteParameterSetBuilder withClientContext(@Nullable final String val) {
            this.clientContext = val;
            return this;
        }
        /**
         * Instanciates a new CallUnmuteParameterSetBuilder
         */
        @Nullable
        protected CallUnmuteParameterSetBuilder(){}
        /**
         * Buils the resulting body object to be passed to the request
         * @return the body object to pass to the request
         */
        @Nonnull
        public CallUnmuteParameterSet build() {
            return new CallUnmuteParameterSet(this);
        }
    }
    /**
     * Gets the functions options from the properties that have been set
     * @return a list of function options for the request
     */
    @Nonnull
    public java.util.List<com.microsoft.graph.options.FunctionOption> getFunctionOptions() {
        final ArrayList<com.microsoft.graph.options.FunctionOption> result = new ArrayList<>();
        if(this.clientContext != null) {
            result.add(new com.microsoft.graph.options.FunctionOption("clientContext", clientContext));
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
