// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models.extensions;
import com.microsoft.graph.concurrency.*;
import com.microsoft.graph.core.*;
import com.microsoft.graph.http.*;
import com.microsoft.graph.options.*;
import com.microsoft.graph.serializer.*;
import java.util.Arrays;
import java.util.EnumSet;


import com.google.gson.JsonObject;
import com.google.gson.JsonElement;
import com.google.gson.annotations.*;
import java.util.HashMap;
import java.util.Map;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Conditional Access Users.
 */
public class ConditionalAccessUsers implements IJsonBackedObject {

    @SerializedName("@odata.type")
    @Expose
    public String oDataType;

    private transient AdditionalDataManager additionalDataManager = new AdditionalDataManager(this);

    @Override
    public final AdditionalDataManager additionalDataManager() {
        return additionalDataManager;
    }

    /**
     * The Include Users.
     * User IDs in scope of policy unless explicitly excluded, or None or All or GuestsOrExternalUsers.
     */
    @SerializedName("includeUsers")
    @Expose
    public java.util.List<String> includeUsers;

    /**
     * The Exclude Users.
     * User IDs excluded from scope of policy and/or GuestsOrExternalUsers.
     */
    @SerializedName("excludeUsers")
    @Expose
    public java.util.List<String> excludeUsers;

    /**
     * The Include Groups.
     * Group IDs in scope of policy unless explicitly excluded, or All.
     */
    @SerializedName("includeGroups")
    @Expose
    public java.util.List<String> includeGroups;

    /**
     * The Exclude Groups.
     * Group IDs excluded from scope of policy.
     */
    @SerializedName("excludeGroups")
    @Expose
    public java.util.List<String> excludeGroups;

    /**
     * The Include Roles.
     * Role IDs in scope of policy unless explicitly excluded, or All.
     */
    @SerializedName("includeRoles")
    @Expose
    public java.util.List<String> includeRoles;

    /**
     * The Exclude Roles.
     * Role IDs excluded from scope of policy.
     */
    @SerializedName("excludeRoles")
    @Expose
    public java.util.List<String> excludeRoles;


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
    public JsonObject getRawObject() {
        return rawObject;
    }

    /**
     * Gets serializer
     *
     * @return the serializer
     */
    protected ISerializer getSerializer() {
        return serializer;
    }

    /**
     * Sets the raw JSON object
     *
     * @param serializer the serializer
     * @param json the JSON object to set this object to
     */
    public void setRawObject(final ISerializer serializer, final JsonObject json) {
        this.serializer = serializer;
        rawObject = json;

    }
}
