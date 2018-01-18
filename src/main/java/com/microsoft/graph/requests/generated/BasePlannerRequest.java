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
 * The class for the Base Planner Request.
 */
public class BasePlannerRequest extends BaseRequest implements IBasePlannerRequest {

    /**
     * The request for the Planner
     *
     * @param requestUrl The request url
     * @param client The service client
     * @param requestOptions The options for this request
     * @param responseClass The class of the reponse
     */
    public BasePlannerRequest(final String requestUrl,
            final IBaseClient client,
            final java.util.List<? extends Option> requestOptions,
            final Class responseClass) {
        super(requestUrl, client, requestOptions, responseClass);
    }

    /**
     * Gets the Planner from the service
     * @param callback The callback to be called after success or failure.
     */
    public void get(final ICallback<Planner> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the Planner from the service
     * @return The Planner from the request.
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public Planner get() throws ClientException {
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
     * Patches this Planner with a source
     * @param sourcePlanner The source object with updates
     * @param callback The callback to be called after success or failure.
     */
    public void patch(final Planner sourcePlanner, final ICallback<Planner> callback) {
        send(HttpMethod.PATCH, callback, sourcePlanner);
    }

    /**
     * Patches this Planner with a source
     * @param sourcePlanner The source object with updates
     * @return The updated Planner
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public Planner patch(final Planner sourcePlanner) throws ClientException {
        return send(HttpMethod.PATCH, sourcePlanner);
    }

    /**
     * Creates a Planner with a new object
     * @param newPlanner The new object to create
     * @param callback The callback to be called after success or failure.
     */
    public void post(final Planner newPlanner, final ICallback<Planner> callback) {
        send(HttpMethod.POST, callback, newPlanner);
    }

    /**
     * Creates a Planner with a new object
     * @param newPlanner The new object to create
     * @return The created Planner
     * @throws ClientException This exception occurs if the request was unable to complete for any reason.
     */
    public Planner post(final Planner newPlanner) throws ClientException {
        return send(HttpMethod.POST, newPlanner);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value The select clause
     * @return The updated request
     */
     public IPlannerRequest select(final String value) {
         getQueryOptions().add(new QueryOption("$select", value));
         return (PlannerRequest)this;
     }

    /**
     * Sets the expand clause for the request
     *
     * @param value The expand clause
     * @return The updated request
     */
     public IPlannerRequest expand(final String value) {
         getQueryOptions().add(new QueryOption("$expand", value));
         return (PlannerRequest)this;
     }

}

