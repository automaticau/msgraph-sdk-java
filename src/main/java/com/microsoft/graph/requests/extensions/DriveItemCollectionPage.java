// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.models.extensions.DriveItem;
import com.microsoft.graph.requests.extensions.DriveItemCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.extensions.DriveItemCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Drive Item Collection Page.
 */
public class DriveItemCollectionPage extends BaseCollectionPage<DriveItem, DriveItemCollectionRequestBuilder> {

    /**
     * A collection page for DriveItem
     *
     * @param response the serialized DriveItemCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public DriveItemCollectionPage(@Nonnull final DriveItemCollectionResponse response, @Nonnull final DriveItemCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for DriveItem
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public DriveItemCollectionPage(@Nonnull final java.util.List<DriveItem> pageContents, @Nullable final DriveItemCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
