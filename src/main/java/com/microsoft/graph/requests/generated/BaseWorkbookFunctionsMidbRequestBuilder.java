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
 * The class for the Base Workbook Functions Midb Request Builder.
 */
public class BaseWorkbookFunctionsMidbRequestBuilder extends BaseActionRequestBuilder {

    /**
     * The request builder for this WorkbookFunctionsMidb
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     */
    public BaseWorkbookFunctionsMidbRequestBuilder(final String requestUrl, final IBaseClient client, final java.util.List<? extends Option> requestOptions, final com.google.gson.JsonElement text, final com.google.gson.JsonElement startNum, final com.google.gson.JsonElement numBytes) {
        super(requestUrl, client, requestOptions);
        bodyParams.put("text", text);
        bodyParams.put("startNum", startNum);
        bodyParams.put("numBytes", numBytes);
    }

    /**
     * Creates the IWorkbookFunctionsMidbRequest
     *
     * @return The IWorkbookFunctionsMidbRequest instance
     */
    public IWorkbookFunctionsMidbRequest buildRequest() {
        return buildRequest(getOptions());
    }

    /**
     * Creates the IWorkbookFunctionsMidbRequest with specific requestOptions instead of the existing requestOptions
     *
     * @param requestOptions the options for the request
     * @return The IWorkbookFunctionsMidbRequest instance
     */
    public IWorkbookFunctionsMidbRequest buildRequest(final java.util.List<? extends Option> requestOptions) {
        WorkbookFunctionsMidbRequest request = new WorkbookFunctionsMidbRequest(
                getRequestUrl(),
                getClient(),
                requestOptions
        );

        if (hasParameter("text")) {
            request.body.text = getParameter("text");
        }

        if (hasParameter("startNum")) {
            request.body.startNum = getParameter("startNum");
        }

        if (hasParameter("numBytes")) {
            request.body.numBytes = getParameter("numBytes");
        }

        return request;
    }
}
