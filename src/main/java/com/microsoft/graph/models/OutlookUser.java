// Template Source: BaseEntity.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models;
import com.microsoft.graph.serializer.ISerializer;
import com.microsoft.graph.serializer.IJsonBackedObject;
import com.microsoft.graph.serializer.AdditionalDataManager;
import java.util.EnumSet;
import com.microsoft.graph.http.BaseCollectionPage;
import com.microsoft.graph.models.OutlookCategory;
import com.microsoft.graph.models.Entity;
import com.microsoft.graph.requests.OutlookCategoryCollectionPage;


import com.google.gson.JsonObject;
import com.google.gson.annotations.SerializedName;
import com.google.gson.annotations.Expose;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Outlook User.
 */
public class OutlookUser extends Entity implements IJsonBackedObject {


    /**
     * The Master Categories.
     * A list of categories defined for the user.
     */
    @SerializedName(value = "masterCategories", alternate = {"MasterCategories"})
    @Expose
	@Nullable
    public OutlookCategoryCollectionPage masterCategories;


    /**
     * Sets the raw JSON object
     *
     * @param serializer the serializer
     * @param json the JSON object to set this object to
     */
    public void setRawObject(@Nonnull final ISerializer serializer, @Nonnull final JsonObject json) {


        if (json.has("masterCategories")) {
            masterCategories = serializer.deserializeObject(json.get("masterCategories"), OutlookCategoryCollectionPage.class);
        }
    }
}
