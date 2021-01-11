// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.models.extensions.ProfilePhoto;
import com.microsoft.graph.requests.extensions.ProfilePhotoCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.extensions.ProfilePhotoCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Profile Photo Collection Page.
 */
public class ProfilePhotoCollectionPage extends BaseCollectionPage<ProfilePhoto, ProfilePhotoCollectionRequestBuilder> {

    /**
     * A collection page for ProfilePhoto
     *
     * @param response the serialized ProfilePhotoCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public ProfilePhotoCollectionPage(@Nonnull final ProfilePhotoCollectionResponse response, @Nonnull final ProfilePhotoCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for ProfilePhoto
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public ProfilePhotoCollectionPage(@Nonnull final java.util.List<ProfilePhoto> pageContents, @Nullable final ProfilePhotoCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
