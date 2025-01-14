// Template Source: BaseEntityCollectionPage.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests;
import com.microsoft.graph.models.ColumnDefinition;
import com.microsoft.graph.requests.ColumnDefinitionCollectionRequestBuilder;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.requests.ColumnDefinitionCollectionResponse;
import com.microsoft.graph.http.BaseCollectionPage;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Column Definition Collection Page.
 */
public class ColumnDefinitionCollectionPage extends BaseCollectionPage<ColumnDefinition, ColumnDefinitionCollectionRequestBuilder> {

    /**
     * A collection page for ColumnDefinition
     *
     * @param response the serialized ColumnDefinitionCollectionResponse from the service
     * @param builder  the request builder for the next collection page
     */
    public ColumnDefinitionCollectionPage(@Nonnull final ColumnDefinitionCollectionResponse response, @Nonnull final ColumnDefinitionCollectionRequestBuilder builder) {
        super(response, builder);
    }

    /**
     * Creates the collection page for ColumnDefinition
     *
     * @param pageContents       the contents of this page
     * @param nextRequestBuilder the request builder for the next page
     */
    public ColumnDefinitionCollectionPage(@Nonnull final java.util.List<ColumnDefinition> pageContents, @Nullable final ColumnDefinitionCollectionRequestBuilder nextRequestBuilder) {
        super(pageContents, nextRequestBuilder);
    }
}
