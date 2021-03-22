// Template Source: BaseEntity.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.models;
import com.microsoft.graph.serializer.ISerializer;
import com.microsoft.graph.serializer.IJsonBackedObject;
import com.microsoft.graph.serializer.AdditionalDataManager;
import java.util.EnumSet;
import com.microsoft.graph.models.IdentitySet;
import com.microsoft.graph.models.File;
import com.microsoft.graph.models.FileSystemInfo;
import com.microsoft.graph.models.Folder;
import com.microsoft.graph.models.Image;
import com.microsoft.graph.models.ItemReference;
import com.microsoft.graph.models.Shared;
import com.microsoft.graph.models.SharepointIds;
import com.microsoft.graph.models.SpecialFolder;
import com.microsoft.graph.models.Video;


import com.google.gson.JsonObject;
import com.google.gson.annotations.SerializedName;
import com.google.gson.annotations.Expose;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Remote Item.
 */
public class RemoteItem implements IJsonBackedObject {

    /** the OData type of the object as returned by the service */
    @SerializedName("@odata.type")
    @Expose
    @Nullable
    public String oDataType;

    private transient AdditionalDataManager additionalDataManager = new AdditionalDataManager(this);

    @Override
    @Nonnull
    public final AdditionalDataManager additionalDataManager() {
        return additionalDataManager;
    }

    /**
     * The Created By.
     * Identity of the user, device, and application which created the item. Read-only.
     */
    @SerializedName(value = "createdBy", alternate = {"CreatedBy"})
    @Expose
	@Nullable
    public IdentitySet createdBy;

    /**
     * The Created Date Time.
     * Date and time of item creation. Read-only.
     */
    @SerializedName(value = "createdDateTime", alternate = {"CreatedDateTime"})
    @Expose
	@Nullable
    public java.time.OffsetDateTime createdDateTime;

    /**
     * The File.
     * Indicates that the remote item is a file. Read-only.
     */
    @SerializedName(value = "file", alternate = {"File"})
    @Expose
	@Nullable
    public File file;

    /**
     * The File System Info.
     * Information about the remote item from the local file system. Read-only.
     */
    @SerializedName(value = "fileSystemInfo", alternate = {"FileSystemInfo"})
    @Expose
	@Nullable
    public FileSystemInfo fileSystemInfo;

    /**
     * The Folder.
     * Indicates that the remote item is a folder. Read-only.
     */
    @SerializedName(value = "folder", alternate = {"Folder"})
    @Expose
	@Nullable
    public Folder folder;

    /**
     * The Id.
     * Unique identifier for the remote item in its drive. Read-only.
     */
    @SerializedName(value = "id", alternate = {"Id"})
    @Expose
	@Nullable
    public String id;

    /**
     * The Image.
     * Image metadata, if the item is an image. Read-only.
     */
    @SerializedName(value = "image", alternate = {"Image"})
    @Expose
	@Nullable
    public Image image;

    /**
     * The Last Modified By.
     * Identity of the user, device, and application which last modified the item. Read-only.
     */
    @SerializedName(value = "lastModifiedBy", alternate = {"LastModifiedBy"})
    @Expose
	@Nullable
    public IdentitySet lastModifiedBy;

    /**
     * The Last Modified Date Time.
     * Date and time the item was last modified. Read-only.
     */
    @SerializedName(value = "lastModifiedDateTime", alternate = {"LastModifiedDateTime"})
    @Expose
	@Nullable
    public java.time.OffsetDateTime lastModifiedDateTime;

    /**
     * The Name.
     * Optional. Filename of the remote item. Read-only.
     */
    @SerializedName(value = "name", alternate = {"Name"})
    @Expose
	@Nullable
    public String name;

    /**
     * The Package.
     * If present, indicates that this item is a package instead of a folder or file. Packages are treated like files in some contexts and folders in others. Read-only.
     */
    @SerializedName(value = "package", alternate = {"Package"})
    @Expose
	@Nullable
    public com.microsoft.graph.models.Package msgraphPackage;

    /**
     * The Parent Reference.
     * Properties of the parent of the remote item. Read-only.
     */
    @SerializedName(value = "parentReference", alternate = {"ParentReference"})
    @Expose
	@Nullable
    public ItemReference parentReference;

    /**
     * The Shared.
     * Indicates that the item has been shared with others and provides information about the shared state of the item. Read-only.
     */
    @SerializedName(value = "shared", alternate = {"Shared"})
    @Expose
	@Nullable
    public Shared shared;

    /**
     * The Sharepoint Ids.
     * Provides interop between items in OneDrive for Business and SharePoint with the full set of item identifiers. Read-only.
     */
    @SerializedName(value = "sharepointIds", alternate = {"SharepointIds"})
    @Expose
	@Nullable
    public SharepointIds sharepointIds;

    /**
     * The Size.
     * Size of the remote item. Read-only.
     */
    @SerializedName(value = "size", alternate = {"Size"})
    @Expose
	@Nullable
    public Long size;

    /**
     * The Special Folder.
     * If the current item is also available as a special folder, this facet is returned. Read-only.
     */
    @SerializedName(value = "specialFolder", alternate = {"SpecialFolder"})
    @Expose
	@Nullable
    public SpecialFolder specialFolder;

    /**
     * The Video.
     * Video metadata, if the item is a video. Read-only.
     */
    @SerializedName(value = "video", alternate = {"Video"})
    @Expose
	@Nullable
    public Video video;

    /**
     * The Web Dav Url.
     * DAV compatible URL for the item.
     */
    @SerializedName(value = "webDavUrl", alternate = {"WebDavUrl"})
    @Expose
	@Nullable
    public String webDavUrl;

    /**
     * The Web Url.
     * URL that displays the resource in the browser. Read-only.
     */
    @SerializedName(value = "webUrl", alternate = {"WebUrl"})
    @Expose
	@Nullable
    public String webUrl;


    /**
     * Sets the raw JSON object
     *
     * @param serializer the serializer
     * @param json the JSON object to set this object to
     */
    public void setRawObject(@Nonnull final ISerializer serializer, @Nonnull final JsonObject json) {

    }
}