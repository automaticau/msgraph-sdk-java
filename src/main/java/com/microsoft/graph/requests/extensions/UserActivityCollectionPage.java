// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.models.extensions.UserActivity;
import com.microsoft.graph.requests.extensions.UserActivityCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.extensions.UserActivityCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the User Activity Collection Page.
 */
public class UserActivityCollectionPage extends BaseCollectionPage<UserActivity, UserActivityCollectionRequestBuilder> {

    /**
     * A collection page for UserActivity
     *
     * @param response the serialized UserActivityCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public UserActivityCollectionPage(@Nonnull final UserActivityCollectionResponse response, @Nonnull final UserActivityCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for UserActivity
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public UserActivityCollectionPage(@Nonnull final java.util.List<UserActivity> pageContents, @Nullable final UserActivityCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
