// Template Source: BaseEntityRequest.java.tt
// ------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.  See License in the project root for license information.
// ------------------------------------------------------------------------------

package com.microsoft.graph.requests.extensions;

import com.microsoft.graph.http.IRequestBuilder;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.concurrency.ICallback;
import com.microsoft.graph.models.extensions.Team;
import com.microsoft.graph.models.generated.TeamVisibilityType;
import com.microsoft.graph.models.generated.ClonableTeamParts;
import java.util.EnumSet;
import com.microsoft.graph.models.extensions.ChatMessage;
import com.microsoft.graph.requests.extensions.ChannelCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.ChannelRequestBuilder;
import com.microsoft.graph.requests.extensions.TeamsAppInstallationCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.TeamsAppInstallationRequestBuilder;
import com.microsoft.graph.requests.extensions.ConversationMemberCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.ConversationMemberRequestBuilder;
import com.microsoft.graph.requests.extensions.TeamsAsyncOperationCollectionRequestBuilder;
import com.microsoft.graph.requests.extensions.TeamsAsyncOperationRequestBuilder;
import com.microsoft.graph.requests.extensions.ScheduleRequestBuilder;
import com.microsoft.graph.requests.extensions.GroupRequestBuilder;
import com.microsoft.graph.requests.extensions.TeamsTemplateRequestBuilder;
import java.util.Arrays;
import java.util.EnumSet;
import javax.annotation.Nullable;
import javax.annotation.Nonnull;
import com.microsoft.graph.core.IBaseClient;
import com.microsoft.graph.http.BaseRequest;
import com.microsoft.graph.http.HttpMethod;

// **NOTE** This file was generated by a tool and any changes will be overwritten.

/**
 * The class for the Team Request.
 */
public class TeamRequest extends BaseRequest<Team> {
	
    /**
     * The request for the Team
     *
     * @param requestUrl     the request URL
     * @param client         the service client
     * @param requestOptions the options for this request
     */
    public TeamRequest(@Nonnull final String requestUrl, @Nonnull final IBaseClient client, @Nullable final java.util.List<? extends com.microsoft.graph.options.Option> requestOptions) {
        super(requestUrl, client, requestOptions, Team.class);
    }

    /**
     * Gets the Team from the service
     *
     * @param callback the callback to be called after success or failure
     */
    public void get(@Nonnull final ICallback<? super Team> callback) {
        send(HttpMethod.GET, callback, null);
    }

    /**
     * Gets the Team from the service
     *
     * @return the Team from the request
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Team get() throws ClientException {
       return send(HttpMethod.GET, null);
    }

    /**
     * Delete this item from the service
     *
     * @param callback the callback when the deletion action has completed
     */
    public void delete(@Nonnull final ICallback<? super Team> callback) {
        send(HttpMethod.DELETE, callback, null);
    }

    /**
     * Delete this item from the service
     *
     * @throws ClientException if there was an exception during the delete operation
     */
    public void delete() throws ClientException {
        send(HttpMethod.DELETE, null);
    }

    /**
     * Patches this Team with a source
     *
     * @param sourceTeam the source object with updates
     * @param callback the callback to be called after success or failure
     */
    public void patch(@Nonnull final Team sourceTeam, @Nonnull final ICallback<? super Team> callback) {
        send(HttpMethod.PATCH, callback, sourceTeam);
    }

    /**
     * Patches this Team with a source
     *
     * @param sourceTeam the source object with updates
     * @return the updated Team
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Team patch(@Nonnull final Team sourceTeam) throws ClientException {
        return send(HttpMethod.PATCH, sourceTeam);
    }

    /**
     * Creates a Team with a new object
     *
     * @param newTeam the new object to create
     * @param callback the callback to be called after success or failure
     */
    public void post(@Nonnull final Team newTeam, @Nonnull final ICallback<? super Team> callback) {
        send(HttpMethod.POST, callback, newTeam);
    }

    /**
     * Creates a Team with a new object
     *
     * @param newTeam the new object to create
     * @return the created Team
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Team post(@Nonnull final Team newTeam) throws ClientException {
        return send(HttpMethod.POST, newTeam);
    }

    /**
     * Creates a Team with a new object
     *
     * @param newTeam the object to create/update
     * @param callback the callback to be called after success or failure
     */
    public void put(@Nonnull final Team newTeam, @Nonnull final ICallback<? super Team> callback) {
        send(HttpMethod.PUT, callback, newTeam);
    }

    /**
     * Creates a Team with a new object
     *
     * @param newTeam the object to create/update
     * @return the created Team
     * @throws ClientException this exception occurs if the request was unable to complete for any reason
     */
    @Nullable
    public Team put(@Nonnull final Team newTeam) throws ClientException {
        return send(HttpMethod.PUT, newTeam);
    }

    /**
     * Sets the select clause for the request
     *
     * @param value the select clause
     * @return the updated request
     */
     @Nonnull
     public TeamRequest select(@Nonnull final String value) {
         addSelectOption(value);
         return this;
     }

    /**
     * Sets the expand clause for the request
     *
     * @param value the expand clause
     * @return the updated request
     */
     @Nonnull
     public TeamRequest expand(@Nonnull final String value) {
         addExpandOption(value);
         return this;
     }

}

