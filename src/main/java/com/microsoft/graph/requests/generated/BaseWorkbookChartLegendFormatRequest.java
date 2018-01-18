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
 * The class for the Base Workbook Chart Legend Format Request.
 */
public class BaseWorkbookChartLegendFormatRequest extends BaseRequest implements IBaseWorkbookChartLegendFormatRequest {

    /**
     * The request for the WorkbookChartLegendFormat
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     * @param responseClass The class of the reponse
     */
    public BaseWorkbookChartLegendFormatRequest(final String requestUrl,
            final IBaseClient client,
            final java.util.List<? extends Option> requestOptions,
            final Class responseClass) {
        super(requestUrl, client, requestOptions, responseClass);
    }

    /**
     * Gets the WorkbookChartLegendFormat from the service
     * @param callback The callback to be called after success or failure.
     */
    public void get(final ICallback<WorkbookChartLegendFormat> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the WorkbookChartLegendFormat from the service
     * @return The WorkbookChartLegendFormat from the request.
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public WorkbookChartLegendFormat get() throws ClientException {
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
     * Patches this WorkbookChartLegendFormat with a source
     * @param sourceWorkbookChartLegendFormat The source object with updates
     * @param callback The callback to be called after success or failure.
     */
    public void patch(final WorkbookChartLegendFormat sourceWorkbookChartLegendFormat, final ICallback<WorkbookChartLegendFormat> callback) {
        send(HttpMethod.PATCH, callback, sourceWorkbookChartLegendFormat);
    }

    /**
     * Patches this WorkbookChartLegendFormat with a source
     * @param sourceWorkbookChartLegendFormat The source object with updates
     * @return The updated WorkbookChartLegendFormat
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public WorkbookChartLegendFormat patch(final WorkbookChartLegendFormat sourceWorkbookChartLegendFormat) throws ClientException {
        return send(HttpMethod.PATCH, sourceWorkbookChartLegendFormat);
    }

    /**
     * Creates a WorkbookChartLegendFormat with a new object
     * @param newWorkbookChartLegendFormat The new object to create
     * @param callback The callback to be called after success or failure.
     */
    public void post(final WorkbookChartLegendFormat newWorkbookChartLegendFormat, final ICallback<WorkbookChartLegendFormat> callback) {
        send(HttpMethod.POST, callback, newWorkbookChartLegendFormat);
    }

    /**
     * Creates a WorkbookChartLegendFormat with a new object
     * @param newWorkbookChartLegendFormat The new object to create
     * @return The created WorkbookChartLegendFormat
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public WorkbookChartLegendFormat post(final WorkbookChartLegendFormat newWorkbookChartLegendFormat) throws ClientException {
        return send(HttpMethod.POST, newWorkbookChartLegendFormat);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value The select clause
     * @return The updated request
     */
     public IWorkbookChartLegendFormatRequest select(final String value) {
         getQueryOptions().add(new QueryOption("$select", value));
         return (WorkbookChartLegendFormatRequest)this;
     }

    /**
     * Sets the expand clause for the request
     *
     * @param value The expand clause
     * @return The updated request
     */
     public IWorkbookChartLegendFormatRequest expand(final String value) {
         getQueryOptions().add(new QueryOption("$expand", value));
         return (WorkbookChartLegendFormatRequest)this;
     }

}

