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
import com.microsoft.graph.models.Alert;
import com.microsoft.graph.models.SecureScoreControlProfile;
import com.microsoft.graph.models.SecureScore;
import com.microsoft.graph.models.Entity;
import com.microsoft.graph.requests.AlertCollectionPage;
import com.microsoft.graph.requests.SecureScoreControlProfileCollectionPage;
import com.microsoft.graph.requests.SecureScoreCollectionPage;


import com.google.gson.JsonObject;
import com.google.gson.annotations.SerializedName;
import com.google.gson.annotations.Expose;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Security.
 */
public class Security extends Entity implements IJsonBackedObject {


    /**
     * The Alerts.
     * Read-only. Nullable.
     */
    @SerializedName(value = "alerts", alternate = {"Alerts"})
    @Expose
	@Nullable
    public AlertCollectionPage alerts;

    /**
     * The Secure Score Control Profiles.
     * 
     */
    @SerializedName(value = "secureScoreControlProfiles", alternate = {"SecureScoreControlProfiles"})
    @Expose
	@Nullable
    public SecureScoreControlProfileCollectionPage secureScoreControlProfiles;

    /**
     * The Secure Scores.
     * 
     */
    @SerializedName(value = "secureScores", alternate = {"SecureScores"})
    @Expose
	@Nullable
    public SecureScoreCollectionPage secureScores;


    /**
     * Sets the raw JSON object
     *
     * @param serializer the serializer
     * @param json the JSON object to set this object to
     */
    public void setRawObject(@Nonnull final ISerializer serializer, @Nonnull final JsonObject json) {


        if (json.has("alerts")) {
            alerts = serializer.deserializeObject(json.get("alerts"), AlertCollectionPage.class);
        }

        if (json.has("secureScoreControlProfiles")) {
            secureScoreControlProfiles = serializer.deserializeObject(json.get("secureScoreControlProfiles"), SecureScoreControlProfileCollectionPage.class);
        }

        if (json.has("secureScores")) {
            secureScores = serializer.deserializeObject(json.get("secureScores"), SecureScoreCollectionPage.class);
        }
    }
}
