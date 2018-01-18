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
 * The class for the Base Workbook Table Row Count Request.
 */
public class BaseWorkbookTableRowCountRequest extends BaseRequest implements IBaseWorkbookTableRowCountRequest {

    /**
     * The request for this WorkbookTableRowCount
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     */
    public BaseWorkbookTableRowCountRequest(final String requestUrl, final IBaseClient client, final java.util.List<? extends Option> requestOptions) {
        super(requestUrl, client, requestOptions, Integer.class);
    }

    /**
     * Gets the Integer
     *
     * @param callback The callback to be called after success or failure.
     */
    public void get(final ICallback<Integer> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the Integer
     *
     * @return The Integer
     * @throws ClientException An exception occurs if there was an error while the request was sent.
     */
    public Integer get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value The select clause
     * @return The updated request
     */
    public IWorkbookTableRowCountRequest select(final String value) {
        getQueryOptions().add(new QueryOption("$select", value));
        return (WorkbookTableRowCountRequest)this;
    }

    /**
     * Sets the expand clause for the request
     *
     * @param value The expand clause
     * @return The updated request
     */
    public IWorkbookTableRowCountRequest expand(final String value) {
        getQueryOptions().add(new QueryOption("$expand", value));
        return (WorkbookTableRowCountRequest)this;
    }

}
