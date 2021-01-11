// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;
import com.microsoft.graph.models.extensions.ManagedDeviceMobileAppConfigurationAssignment;
import com.microsoft.graph.requests.extensions.ManagedDeviceMobileAppConfigurationAssignmentCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.extensions.ManagedDeviceMobileAppConfigurationAssignmentCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Managed Device Mobile App Configuration Assignment Collection Page.
 */
public class ManagedDeviceMobileAppConfigurationAssignmentCollectionPage extends BaseCollectionPage<ManagedDeviceMobileAppConfigurationAssignment, ManagedDeviceMobileAppConfigurationAssignmentCollectionRequestBuilder> {

    /**
     * A collection page for ManagedDeviceMobileAppConfigurationAssignment
     *
     * @param response the serialized ManagedDeviceMobileAppConfigurationAssignmentCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public ManagedDeviceMobileAppConfigurationAssignmentCollectionPage(@Nonnull final ManagedDeviceMobileAppConfigurationAssignmentCollectionResponse response, @Nonnull final ManagedDeviceMobileAppConfigurationAssignmentCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for ManagedDeviceMobileAppConfigurationAssignment
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public ManagedDeviceMobileAppConfigurationAssignmentCollectionPage(@Nonnull final java.util.List<ManagedDeviceMobileAppConfigurationAssignment> pageContents, @Nullable final ManagedDeviceMobileAppConfigurationAssignmentCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
