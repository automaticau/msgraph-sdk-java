// Template Source: BaseEntity.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models.extensions;
import com.microsoft.graph.serializer.ISerializer;
import com.microsoft.graph.serializer.IJsonBackedObject;
import com.microsoft.graph.serializer.AdditionalDataManager;
import java.util.EnumSet;
import com.microsoft.graph.models.generated.InferenceClassificationType;
import com.microsoft.graph.models.extensions.EmailAddress;
import com.microsoft.graph.models.extensions.Entity;


import com.google.gson.JsonObject;
import com.google.gson.annotations.SerializedName;
import com.google.gson.annotations.Expose;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Inference Classification Override.
 */
public class InferenceClassificationOverride extends Entity implements IJsonBackedObject {


    /**
     * The Classify As.
     * Specifies how incoming messages from a specific sender should always be classified as. Possible values are: focused, other.
     */
    @SerializedName(value = "classifyAs", alternate = {"ClassifyAs"})
    @Expose
    public InferenceClassificationType classifyAs;

    /**
     * The Sender Email Address.
     * The email address information of the sender for whom the override is created.
     */
    @SerializedName(value = "senderEmailAddress", alternate = {"SenderEmailAddress"})
    @Expose
    public EmailAddress senderEmailAddress;


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
