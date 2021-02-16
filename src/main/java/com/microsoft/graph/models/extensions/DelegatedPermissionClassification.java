// Template Source: BaseEntity.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models.extensions;
import com.microsoft.graph.serializer.ISerializer;
import com.microsoft.graph.serializer.IJsonBackedObject;
import com.microsoft.graph.serializer.AdditionalDataManager;
import java.util.EnumSet;
import com.microsoft.graph.models.generated.PermissionClassificationType;
import com.microsoft.graph.models.extensions.Entity;


import com.google.gson.JsonObject;
import com.google.gson.annotations.SerializedName;
import com.google.gson.annotations.Expose;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Delegated Permission Classification.
 */
public class DelegatedPermissionClassification extends Entity implements IJsonBackedObject {


    /**
     * The Classification.
     * The classification value being given. Possible value: low. Does not support $filter.
     */
    @SerializedName(value = "classification", alternate = {"Classification"})
    @Expose
    public PermissionClassificationType classification;

    /**
     * The Permission Id.
     * The unique identifier (id) for the delegated permission listed in the publishedPermissionScopes collection of the servicePrincipal. Required on create. Does not support $filter.
     */
    @SerializedName(value = "permissionId", alternate = {"PermissionId"})
    @Expose
    public String permissionId;

    /**
     * The Permission Name.
     * The claim value (value) for the delegated permission listed in the publishedPermissionScopes collection of the servicePrincipal. Does not support $filter.
     */
    @SerializedName(value = "permissionName", alternate = {"PermissionName"})
    @Expose
    public String permissionName;


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
