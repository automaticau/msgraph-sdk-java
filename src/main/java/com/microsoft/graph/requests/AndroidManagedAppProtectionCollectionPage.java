// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;
import com.microsoft.graph.models.AndroidManagedAppProtection;
import com.microsoft.graph.requests.AndroidManagedAppProtectionCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.AndroidManagedAppProtectionCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Android Managed App Protection Collection Page.
 */
public class AndroidManagedAppProtectionCollectionPage extends BaseCollectionPage<AndroidManagedAppProtection, AndroidManagedAppProtectionCollectionRequestBuilder> {

    /**
     * A collection page for AndroidManagedAppProtection
     *
     * @param response the serialized AndroidManagedAppProtectionCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public AndroidManagedAppProtectionCollectionPage(@Nonnull final AndroidManagedAppProtectionCollectionResponse response, @Nonnull final AndroidManagedAppProtectionCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for AndroidManagedAppProtection
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public AndroidManagedAppProtectionCollectionPage(@Nonnull final java.util.List<AndroidManagedAppProtection> pageContents, @Nullable final AndroidManagedAppProtectionCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
