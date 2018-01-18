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
 * The class for the Base Workbook Close Session Request Builder.
 */
public class BaseWorkbookCloseSessionRequestBuilder extends BaseActionRequestBuilder {

    /**
     * The request builder for this WorkbookCloseSession
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     */
    public BaseWorkbookCloseSessionRequestBuilder(final String requestUrl, final IBaseClient client, final java.util.List<? extends Option> requestOptions) {
        super(requestUrl, client, requestOptions);
    }

    /**
     * Creates the IWorkbookCloseSessionRequest
     *
     * @return The IWorkbookCloseSessionRequest instance
     */
    public IWorkbookCloseSessionRequest buildRequest() {
        return buildRequest(getOptions());
    }

    /**
     * Creates the IWorkbookCloseSessionRequest with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for the request
     * @return The IWorkbookCloseSessionRequest instance
     */
    public IWorkbookCloseSessionRequest buildRequest(final java.util.List<? extends Option> requestOptions) {
        WorkbookCloseSessionRequest request = new WorkbookCloseSessionRequest(
                getRequestUrl(),
                getClient(),
                requestOptions
        );

        return request;
    }
}
