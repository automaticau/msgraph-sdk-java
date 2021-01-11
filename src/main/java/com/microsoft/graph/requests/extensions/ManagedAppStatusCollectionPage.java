// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.models.extensions.ManagedAppStatus;
import com.microsoft.graph.requests.extensions.ManagedAppStatusCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.extensions.ManagedAppStatusCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Managed App Status Collection Page.
 */
public class ManagedAppStatusCollectionPage extends BaseCollectionPage<ManagedAppStatus, ManagedAppStatusCollectionRequestBuilder> {

    /**
     * A collection page for ManagedAppStatus
     *
     * @param response the serialized ManagedAppStatusCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public ManagedAppStatusCollectionPage(@Nonnull final ManagedAppStatusCollectionResponse response, @Nonnull final ManagedAppStatusCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for ManagedAppStatus
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public ManagedAppStatusCollectionPage(@Nonnull final java.util.List<ManagedAppStatus> pageContents, @Nullable final ManagedAppStatusCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
