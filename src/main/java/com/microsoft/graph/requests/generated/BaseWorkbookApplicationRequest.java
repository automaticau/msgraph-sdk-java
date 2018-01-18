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
 * The class for the Base Workbook Application Request.
 */
public class BaseWorkbookApplicationRequest extends BaseRequest implements IBaseWorkbookApplicationRequest {

    /**
     * The request for the WorkbookApplication
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     * @param responseClass The class of the reponse
     */
    public BaseWorkbookApplicationRequest(final String requestUrl,
            final IBaseClient client,
            final java.util.List<? extends Option> requestOptions,
            final Class responseClass) {
        super(requestUrl, client, requestOptions, responseClass);
    }

    /**
     * Gets the WorkbookApplication from the service
     * @param callback The callback to be called after success or failure.
     */
    public void get(final ICallback<WorkbookApplication> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the WorkbookApplication from the service
     * @return The WorkbookApplication from the request.
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public WorkbookApplication get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service.
     * @param callback The callback when the deletion action has completed
     */
    public void delete(final ICallback<Void> callback) {{
        send(HttpMethod.DELETE, callback, null);
    }}

    /**
     * Delete this item from the service.
     * @throws ClientException if there was an exception during the delete operation
     */
    public void delete() throws ClientException {{
        send(HttpMethod.DELETE, null);
    }}

    /**
     * Patches this WorkbookApplication with a source
     * @param sourceWorkbookApplication The source object with updates
     * @param callback The callback to be called after success or failure.
     */
    public void patch(final WorkbookApplication sourceWorkbookApplication, final ICallback<WorkbookApplication> callback) {
        send(HttpMethod.PATCH, callback, sourceWorkbookApplication);
    }

    /**
     * Patches this WorkbookApplication with a source
     * @param sourceWorkbookApplication The source object with updates
     * @return The updated WorkbookApplication
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public WorkbookApplication patch(final WorkbookApplication sourceWorkbookApplication) throws ClientException {
        return send(HttpMethod.PATCH, sourceWorkbookApplication);
    }

    /**
     * Creates a WorkbookApplication with a new object
     * @param newWorkbookApplication The new object to create
     * @param callback The callback to be called after success or failure.
     */
    public void post(final WorkbookApplication newWorkbookApplication, final ICallback<WorkbookApplication> callback) {
        send(HttpMethod.POST, callback, newWorkbookApplication);
    }

    /**
     * Creates a WorkbookApplication with a new object
     * @param newWorkbookApplication The new object to create
     * @return The created WorkbookApplication
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public WorkbookApplication post(final WorkbookApplication newWorkbookApplication) throws ClientException {
        return send(HttpMethod.POST, newWorkbookApplication);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value The select clause
     * @return The updated request
     */
     public IWorkbookApplicationRequest select(final String value) {
         getQueryOptions().add(new QueryOption("$select", value));
         return (WorkbookApplicationRequest)this;
     }

    /**
     * Sets the expand clause for the request
     *
     * @param value The expand clause
     * @return The updated request
     */
     public IWorkbookApplicationRequest expand(final String value) {
         getQueryOptions().add(new QueryOption("$expand", value));
         return (WorkbookApplicationRequest)this;
     }

}

