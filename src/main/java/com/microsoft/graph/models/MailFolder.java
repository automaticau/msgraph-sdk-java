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
import com.microsoft.graph.models.MailFolder;
import com.microsoft.graph.models.MessageRule;
import com.microsoft.graph.models.Message;
import com.microsoft.graph.models.MultiValueLegacyExtendedProperty;
import com.microsoft.graph.models.SingleValueLegacyExtendedProperty;
import com.microsoft.graph.models.Entity;
import com.microsoft.graph.requests.MailFolderCollectionPage;
import com.microsoft.graph.requests.MessageRuleCollectionPage;
import com.microsoft.graph.requests.MessageCollectionPage;
import com.microsoft.graph.requests.MultiValueLegacyExtendedPropertyCollectionPage;
import com.microsoft.graph.requests.SingleValueLegacyExtendedPropertyCollectionPage;


import com.google.gson.JsonObject;
import com.google.gson.annotations.SerializedName;
import com.google.gson.annotations.Expose;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Mail Folder.
 */
public class MailFolder extends Entity implements IJsonBackedObject {


    /**
     * The Child Folder Count.
     * The number of immediate child mailFolders in the current mailFolder.
     */
    @SerializedName(value = "childFolderCount", alternate = {"ChildFolderCount"})
    @Expose
	@Nullable
    public Integer childFolderCount;

    /**
     * The Display Name.
     * The mailFolder's display name.
     */
    @SerializedName(value = "displayName", alternate = {"DisplayName"})
    @Expose
	@Nullable
    public String displayName;

    /**
     * The Parent Folder Id.
     * The unique identifier for the mailFolder's parent mailFolder.
     */
    @SerializedName(value = "parentFolderId", alternate = {"ParentFolderId"})
    @Expose
	@Nullable
    public String parentFolderId;

    /**
     * The Total Item Count.
     * The number of items in the mailFolder.
     */
    @SerializedName(value = "totalItemCount", alternate = {"TotalItemCount"})
    @Expose
	@Nullable
    public Integer totalItemCount;

    /**
     * The Unread Item Count.
     * The number of items in the mailFolder marked as unread.
     */
    @SerializedName(value = "unreadItemCount", alternate = {"UnreadItemCount"})
    @Expose
	@Nullable
    public Integer unreadItemCount;

    /**
     * The Child Folders.
     * The collection of child folders in the mailFolder.
     */
    @SerializedName(value = "childFolders", alternate = {"ChildFolders"})
    @Expose
	@Nullable
    public MailFolderCollectionPage childFolders;

    /**
     * The Message Rules.
     * The collection of rules that apply to the user's Inbox folder.
     */
    @SerializedName(value = "messageRules", alternate = {"MessageRules"})
    @Expose
	@Nullable
    public MessageRuleCollectionPage messageRules;

    /**
     * The Messages.
     * The collection of messages in the mailFolder.
     */
    @SerializedName(value = "messages", alternate = {"Messages"})
    @Expose
	@Nullable
    public MessageCollectionPage messages;

    /**
     * The Multi Value Extended Properties.
     * The collection of multi-value extended properties defined for the mailFolder. Read-only. Nullable.
     */
    @SerializedName(value = "multiValueExtendedProperties", alternate = {"MultiValueExtendedProperties"})
    @Expose
	@Nullable
    public MultiValueLegacyExtendedPropertyCollectionPage multiValueExtendedProperties;

    /**
     * The Single Value Extended Properties.
     * The collection of single-value extended properties defined for the mailFolder. Read-only. Nullable.
     */
    @SerializedName(value = "singleValueExtendedProperties", alternate = {"SingleValueExtendedProperties"})
    @Expose
	@Nullable
    public SingleValueLegacyExtendedPropertyCollectionPage singleValueExtendedProperties;


    /**
     * Sets the raw JSON object
     *
     * @param serializer the serializer
     * @param json the JSON object to set this object to
     */
    public void setRawObject(@Nonnull final ISerializer serializer, @Nonnull final JsonObject json) {


        if (json.has("childFolders")) {
            childFolders = serializer.deserializeObject(json.get("childFolders"), MailFolderCollectionPage.class);
        }

        if (json.has("messageRules")) {
            messageRules = serializer.deserializeObject(json.get("messageRules"), MessageRuleCollectionPage.class);
        }

        if (json.has("messages")) {
            messages = serializer.deserializeObject(json.get("messages"), MessageCollectionPage.class);
        }

        if (json.has("multiValueExtendedProperties")) {
            multiValueExtendedProperties = serializer.deserializeObject(json.get("multiValueExtendedProperties"), MultiValueLegacyExtendedPropertyCollectionPage.class);
        }

        if (json.has("singleValueExtendedProperties")) {
            singleValueExtendedProperties = serializer.deserializeObject(json.get("singleValueExtendedProperties"), SingleValueLegacyExtendedPropertyCollectionPage.class);
        }
    }
}
