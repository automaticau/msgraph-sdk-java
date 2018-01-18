// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.generated;

import com.microsoft.graph.concurrency.*;
import com.microsoft.graph.core.*;
import com.microsoft.graph.models.extensions.*;
import com.microsoft.graph.models.generated.*;
import com.microsoft.graph.http.*;
import com.microsoft.graph.requests.extensions.*;
import com.microsoft.graph.requests.generated.*;
import com.microsoft.graph.options.*;
import com.microsoft.graph.serializer.*;

import java.util.Arrays;
import java.util.EnumSet;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Base List Request Builder.
 */
public class BaseListRequestBuilder extends BaseRequestBuilder implements IBaseListRequestBuilder {

    /**
     * The request builder for the List
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     */
    public BaseListRequestBuilder(final String requestUrl, final IBaseClient client, final java.util.List<? extends Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }

    /**
     * Creates the request
     */
    public IListRequest buildRequest() {
        return buildRequest(getOptions());
    }

    /**
     * Creates the request with specific requestOptions instead of the existing requestOptions
     */
    public IListRequest buildRequest(final java.util.List<? extends Option> requestOptions) {
        return new ListRequest(getRequestUrl(), getClient(), requestOptions);
    }

    public IColumnDefinitionCollectionRequestBuilder columns() {
        return new ColumnDefinitionCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("columns"), getClient(), null);
    }

    public IColumnDefinitionRequestBuilder columns(final String id) {
        return new ColumnDefinitionRequestBuilder(getRequestUrlWithAdditionalSegment("columns") + "/" + id, getClient(), null);
    }
    public IContentTypeCollectionRequestBuilder contentTypes() {
        return new ContentTypeCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("contentTypes"), getClient(), null);
    }

    public IContentTypeRequestBuilder contentTypes(final String id) {
        return new ContentTypeRequestBuilder(getRequestUrlWithAdditionalSegment("contentTypes") + "/" + id, getClient(), null);
    }

    /**
     * Gets the request builder for Drive.
     */
    public IDriveRequestBuilder drive() {
        return new DriveRequestBuilder(getRequestUrlWithAdditionalSegment("drive"), getClient(), null);
    }
    public IListItemCollectionRequestBuilder items() {
        return new ListItemCollectionRequestBuilder(getRequestUrlWithAdditionalSegment("items"), getClient(), null);
    }

    public IListItemRequestBuilder items(final String id) {
        return new ListItemRequestBuilder(getRequestUrlWithAdditionalSegment("items") + "/" + id, getClient(), null);
    }
}
