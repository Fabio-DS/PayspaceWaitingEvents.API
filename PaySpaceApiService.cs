using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Payspace.Rest.API;
using PaySpaceWaitingEvents.API.Data;
using Payspace.Rest.API.PaySpace.Venuta.Data.Models.Enums;
using PaySpaceWaitingEvents.API.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Microsoft.OData.Client;
using RestSharp;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Payspace.Rest.API.PaySpace.Venuta.Data.Models.Dto;


namespace PaySpaceWaitingEvents.API.Services
{
    public interface IPaySpaceApiService
    {
        Task<PaySpaceApiResponse> SubmitPayElementsAsync(int legalEntityId, string frequency, string run, List<PayElementEntry> entries);
        List<object> GetAvailableCompanies();
        List<string> GetAvailableFrequencies(int legalEntityId);
        Task<bool> ValidateMappingsAsync(int legalEntityId, string frequency, List<string> payElementIds);
        List<object> GetCompanyRuns(int legalEntityId, string frequency);

        Task<bool> EmployeeExistsAsync(int companyId, string employeeNumber);
        Task<Employee> GetEmployeeAsync(int companyId, string employeeNumber);
        Task<EmployeePosition> GetEmployeePositionAsync(int companyId, string employeeNumber);
        Task<EmployeePayRate> GetEmployeePayRateAsync(int companyId, string employeeNumber);
        Task<PaySpaceApiResponse> SubmitEmployeeDataAsync(int legalEntityId, List<PayElementEntry> entries);
        Task<PaySpaceApiResponse> SubmitEmployeeBankingAsync(int legalEntityId, List<PayElementEntry> entries);
        Task<PaySpaceApiResponse> SubmitEmploymentStatusAsync(int legalEntityId, List<PayElementEntry> entries);
        Task<PaySpaceApiResponse> SubmitPositionDataAsync(int legalEntityId, List<PayElementEntry> entries);
        Task<PaySpaceApiResponse> SubmitPayRateAsync(int companyId, List<PayElementEntry> entries);

        Task<PaySpaceApiResponse> SubmitAllCategoriesAsync(int legalEntityId, string frequency, string run, Dictionary<string, List<PayElementEntry>> entriesByCategory);
    }

    public class PaySpaceApiResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public int TotalEntries { get; set; }
        public int SuccessfulEntries { get; set; }
        public int FailedEntries { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
        public List<PayElementResult> Results { get; set; } = new List<PayElementResult>();
    }

    public class PayElementResult
    {
        public string EmployeeId { get; set; }
        public string PayElementId { get; set; }
        public string ComponentCode { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
    }

    public class PaySpaceApiService : IPaySpaceApiService
    {
        private readonly ILogger<PaySpaceApiService> _logger;
        private readonly IConfiguration _configuration;
        private readonly ApplicationDbContext _dbContext;
        private readonly Api _payspaceApi;

        public PaySpaceApiService(
            ILogger<PaySpaceApiService> logger,
            IConfiguration configuration,
            ApplicationDbContext dbContext)
        {
            _logger = logger;
            _configuration = configuration;
            _dbContext = dbContext;

            string clientId = configuration["PaySpace:ClientId"] ?? "nga";
            string clientSecret = configuration["PaySpace:ClientSecret"] ?? "ace9d85b-5813-4708-9789-d074031b9b0b";
            string username = configuration["PaySpace:Username"] ?? "";
            string password = configuration["PaySpace:Password"] ?? "";

            _payspaceApi = new Api(clientId, clientSecret, username, password);
        }

        private async Task<(bool Success, List<string> Messages)> DirectApiSubmit(int companyId, string frequency, string period, object payload)
        {
            var messages = new List<string>();
            try
            {
                _logger.LogInformation($"Making direct API call for company ID {companyId}");

                string token = TokenProvider.GetToken(Api.AuthUri);

                var client = new RestClient(Api.baseRestUri)
                {
                    Timeout = 30000
                };

                string url = $"{companyId}/EmployeeComponent?frequency={Uri.EscapeDataString(frequency)}&period={Uri.EscapeDataString(period)}";

                var request = new RestRequest(url, Method.POST);
                request.AddHeader("Authorization", $"Bearer {token}");
                request.AddHeader("Content-Type", "application/json");

                string jsonPayload = Newtonsoft.Json.JsonConvert.SerializeObject(payload);
                request.AddParameter("application/json", jsonPayload, ParameterType.RequestBody);

                _logger.LogInformation($"Sending request to URL: {Api.baseRestUri}{url}");
                _logger.LogInformation($"Payload: {jsonPayload}");

                var response = await client.ExecuteAsync(request);

                if (response.IsSuccessful)
                {
                    _logger.LogInformation("API call successful");
                    return (true, messages);
                }
                else
                {
                    _logger.LogError($"API call failed: {response.StatusCode} - {response.Content}");
                    messages.Add($"API error ({response.StatusCode}): {response.Content}");
                    return (false, messages);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Exception in DirectApiSubmit");
                messages.Add($"Exception: {ex.Message}");
                return (false, messages);
            }
        }

        private async Task<(bool Success, List<string> Messages)> DirectApiSubmitWithCode(string companyCode, string frequency, string period, object payload)
        {
            var messages = new List<string>();
            try
            {
                _logger.LogInformation($"Making direct API call with company code {companyCode}");

                var agencyCompanies = _payspaceApi.GetAgencyCompanies();
                var company = agencyCompanies
                    .SelectMany(a => a.companies)
                    .FirstOrDefault(c => c.company_code == companyCode);

                if (company == null)
                {
                    messages.Add($"Could not find company with code {companyCode}");
                    return (false, messages);
                }

                var result = await DirectApiSubmit((int)company.company_id, frequency, period, payload);
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Exception in DirectApiSubmitWithCode");
                messages.Add($"Exception: {ex.Message}");
                return (false, messages);
            }
        }

        public List<object> GetAvailableCompanies()
        {
            try
            {
                var agencyCompanies = _payspaceApi.GetAgencyCompanies();

                var companies = new List<object>();
                foreach (var agency in agencyCompanies)
                {
                    foreach (var company in agency.companies)
                    {
                        companies.Add(new
                        {
                            id = company.company_id,
                            companyCode = company.company_code,
                            companyName = company.company_name
                        });
                    }
                }

                _logger.LogInformation($"Retrieved {companies.Count} companies directly from PaySpace API");
                return companies.OrderBy(c => ((dynamic)c).companyName).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting companies from PaySpace API");
                return new List<object>();
            }
        }

        public List<string> GetAvailableFrequencies(int companyId)
        {
            try
            {
                var frequencies = _payspaceApi.GetAllCompanyFrequencies(companyId);

                if (frequencies == null || !frequencies.Any())
                {
                    _logger.LogWarning($"No frequencies found for company ID {companyId}");
                    return new List<string>();
                }

                _logger.LogInformation($"Retrieved {frequencies.Count} frequencies for company ID {companyId}");
                return frequencies.Select(f => f.Value).OrderBy(f => f).ToList();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting frequencies for company ID {companyId}");
                return new List<string>();
            }
        }

        public List<object> GetCompanyRuns(int companyId, string frequency)
        {
            try
            {
                var runs = _payspaceApi.GetCompanyRuns(companyId.ToString(), frequency, true, false);

                if (runs == null || !runs.Any())
                {
                    _logger.LogWarning($"No runs found for company ID {companyId} with frequency {frequency}");
                    return new List<object>();
                }

                var formattedRuns = runs.Select(r => new {
                    Value = r.Value,
                    Description = $"{r.Description} (Ends: {r.PeriodEndDate:yyyy-MM-dd})",
                    PeriodEndDate = r.PeriodEndDate,
                    OrderNumber = r.OrderNumber
                }).OrderByDescending(r => r.PeriodEndDate).ToList<object>();

                _logger.LogInformation($"Retrieved {formattedRuns.Count} runs for company ID {companyId} with frequency {frequency}");
                return formattedRuns;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting runs for company ID {companyId} and frequency {frequency}");
                return new List<object>();
            }
        }

        public async Task<bool> ValidateMappingsAsync(int legalEntityId, string frequency, List<string> payElementIds)
        {
            try
            {
                var legalEntity = await _dbContext.LegalEntities.FindAsync(legalEntityId);
                if (legalEntity == null)
                {
                    _logger.LogWarning($"Legal entity with ID {legalEntityId} not found");
                    return false;
                }

                foreach (var payElementId in payElementIds)
                {
                    var componentCodes = _payspaceApi.GetComponentCodeLookupDetail(
                        legalEntity.CompanyCode,
                        frequency,
                        payElementId
                    );

                    if (componentCodes == null || !componentCodes.Any())
                    {
                        _logger.LogWarning($"No component code found for pay element {payElementId} in frequency {frequency}");
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error validating mappings for legal entity {legalEntityId} and frequency {frequency}");
                return false;
            }
        }

        public async Task<PaySpaceApiResponse> SubmitPayElementsAsync(int payspaceCompanyId, string frequency, string run, List<PayElementEntry> entries)
        {
            var response = new PaySpaceApiResponse
            {
                Success = false,
                TotalEntries = entries.Count,
                SuccessfulEntries = 0,
                FailedEntries = 0,
                Errors = new List<string>()
            };

            try
            {
                _logger.LogInformation($"Processing for PaySpace company ID {payspaceCompanyId}, frequency {frequency}, run {run}");

                // Validate inputs to prevent null errors
                if (string.IsNullOrEmpty(frequency))
                {
                    response.Errors.Add("Frequency cannot be empty");
                    response.Message = "Invalid parameters";
                    response.FailedEntries = entries.Count;
                    return response;
                }

                if (string.IsNullOrEmpty(run))
                {
                    response.Errors.Add("Run cannot be empty");
                    response.Message = "Invalid parameters";
                    response.FailedEntries = entries.Count;
                    return response;
                }

                if (_payspaceApi == null)
                {
                    response.Message = "PaySpace API client is null";
                    response.FailedEntries = entries.Count;
                    response.Errors.Add(response.Message);
                    return response;
                }

                string period = run;
                _logger.LogInformation($"Using period: {period} and frequency: {frequency}");

                string companyCode = string.Empty;
                var agencyCompanies = _payspaceApi.GetAgencyCompanies();
                var company = agencyCompanies
                    .SelectMany(a => a.companies)
                    .FirstOrDefault(c => c.company_id == payspaceCompanyId);

                if (company != null)
                {
                    companyCode = company.company_code;
                    _logger.LogInformation($"Found company code {companyCode} for ID {payspaceCompanyId}");
                }
                else
                {
                    _logger.LogWarning($"Could not find company code for ID {payspaceCompanyId}");
                }

                var entriesByEmployee = entries.GroupBy(e => e.EmployeeId).ToList();
                _logger.LogInformation($"Processing {entriesByEmployee.Count} employees with {entries.Count} total entries");

                List<string> globalErrors = new List<string>();

                // VALIDATION
                _logger.LogInformation("Starting validation phase - checking all entries before submitting any");

                var allRequests = new List<(string EmployeeId, PayElementEntry Entry, object Payload)>();
                var validationErrors = new List<(string EmployeeId, string PayElementId, string Error)>();

                foreach (var employeeGroup in entriesByEmployee)
                {
                    string employeeId = employeeGroup.Key;
                    var employeeEntries = employeeGroup.ToList();

                    foreach (var entry in employeeEntries)
                    {
                        try
                        {
                            if (string.IsNullOrEmpty(entry.PaySpaceCompCode))
                            {
                                validationErrors.Add((employeeId, entry.PayElementId, "No component code mapped for pay element"));
                                continue;
                            }

                            string inputType = "Amount";
                            decimal inputValue;

                            if (!string.IsNullOrEmpty(entry.UnitType))
                            {
                                if (entry.UnitType.Equals("days", StringComparison.OrdinalIgnoreCase))
                                {
                                    inputType = "Days";
                                    inputValue = entry.NumberOfUnits ?? 0;
                                }
                                else if (entry.UnitType.Equals("hours", StringComparison.OrdinalIgnoreCase))
                                {
                                    inputType = "Hours";
                                    inputValue = entry.NumberOfUnits ?? 0;
                                }
                                else if (entry.UnitType.Equals("units", StringComparison.OrdinalIgnoreCase))
                                {
                                    inputType = "Units";
                                    inputValue = entry.NumberOfUnits ?? 0;
                                }
                                else
                                {
                                    inputValue = entry.Amount ?? 0;
                                }
                            }
                            else
                            {
                                inputValue = entry.Amount ?? 0;
                            }

                            _logger.LogInformation($"For component {entry.PaySpaceCompCode}, using InputType: {inputType}, InputValue: {inputValue}");

                            var payload = new
                            {
                                EmployeeNumber = employeeId,
                                ComponentCode = entry.PaySpaceCompCode,
                                InputValue = inputValue,
                                InputType = inputType,
                                PayslipAction = "Add",
                                Comments = $"Added by API for {entry.PayElementId}"
                            };

                            allRequests.Add((employeeId, entry, payload));
                        }
                        catch (Exception ex)
                        {
                            validationErrors.Add((employeeId, entry.PayElementId, $"Validation error: {ex.Message}"));
                        }
                    }
                }

                // If any validation errors, don't upload anything
                if (validationErrors.Any())
                {
                    _logger.LogWarning($"Validation failed with {validationErrors.Count} errors. Aborting without submitting any entries.");

                    response.Success = false;
                    response.FailedEntries = entries.Count;

                    foreach (var error in validationErrors)
                    {
                        var errorMsg = $"Validation error for employee {error.EmployeeId}, pay element {error.PayElementId}: {error.Error}";
                        response.Errors.Add(errorMsg);

                        response.Results.Add(new PayElementResult
                        {
                            EmployeeId = error.EmployeeId,
                            PayElementId = error.PayElementId,
                            Success = false,
                            ErrorMessage = error.Error
                        });
                    }

                    response.Message = $"Validation failed with {validationErrors.Count} errors. No entries were submitted.";
                    return response;
                }

                // SUBMISSION - if all entries are validated, submit them
                _logger.LogInformation($"All {allRequests.Count} entries passed validation. Starting submission phase.");

                string token = TokenProvider.GetToken(Api.AuthUri);

                foreach (var request in allRequests)
                {
                    var elementResult = new PayElementResult
                    {
                        EmployeeId = request.EmployeeId ?? "Unknown",
                        PayElementId = request.Entry.PayElementId ?? "Unknown",
                        ComponentCode = request.Entry.PaySpaceCompCode ?? "",
                        Success = false
                    };

                    try
                    {
                        var url = $"{payspaceCompanyId}/EditPayslip?frequency={Uri.EscapeDataString(frequency)}&period={Uri.EscapeDataString(period)}";

                        var client = new RestClient(Api.baseRestUri)
                        {
                            Timeout = 30000
                        };

                        var restRequest = new RestRequest(url, Method.POST);
                        restRequest.AddHeader("Authorization", $"Bearer {token}");
                        restRequest.AddHeader("Content-Type", "application/json");

                        string jsonPayload = Newtonsoft.Json.JsonConvert.SerializeObject(request.Payload);
                        restRequest.AddParameter("application/json", jsonPayload, ParameterType.RequestBody);

                        _logger.LogInformation($"Sending request to URL: {Api.baseRestUri}{url}");
                        _logger.LogInformation($"Payload: {jsonPayload}");

                        var apiResponse = await client.ExecuteAsync(restRequest);

                        if (apiResponse.IsSuccessful)
                        {
                            elementResult.Success = true;
                            response.SuccessfulEntries++;
                            _logger.LogInformation($"Successfully added payslip entry for employee {request.EmployeeId}, component {request.Entry.PaySpaceCompCode}");
                        }
                        else
                        {
                            string errorMessage = $"API error ({apiResponse.StatusCode}): {apiResponse.Content}";
                            elementResult.ErrorMessage = errorMessage;
                            response.FailedEntries++;
                            _logger.LogError($"Failed to add payslip entry: {errorMessage}");
                            globalErrors.Add(errorMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        elementResult.ErrorMessage = $"Exception: {ex.Message}";
                        response.FailedEntries++;
                        _logger.LogError(ex, $"Error processing entry for employee {request.EmployeeId}, pay element {request.Entry.PayElementId}");
                        globalErrors.Add($"Exception for {request.EmployeeId}: {ex.Message}");
                    }

                    response.Results.Add(elementResult);
                }

                response.Success = response.FailedEntries == 0 && response.SuccessfulEntries > 0;
                response.Message = response.Success
                    ? $"Successfully processed all {response.SuccessfulEntries} entries"
                    : $"Processed with {response.FailedEntries} failures and {response.SuccessfulEntries} successes";

                response.Errors = globalErrors.Distinct().ToList();

                foreach (var elementResult in response.Results)
                {
                    if (string.IsNullOrEmpty(elementResult.EmployeeId))
                    {
                        elementResult.EmployeeId = "Unknown";
                        _logger.LogWarning("Found PayElementResult with null or empty EmployeeId");
                    }

                    if (string.IsNullOrEmpty(elementResult.PayElementId))
                    {
                        elementResult.PayElementId = "Unknown";
                        _logger.LogWarning("Found PayElementResult with null or empty PayElementId");
                    }
                }

                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error submitting pay elements to PaySpace");

                response.Success = false;
                response.Message = $"Error: {ex.Message}";
                response.FailedEntries = entries.Count - response.SuccessfulEntries;
                response.Errors.Add(ex.Message);

                return response;
            }
        }
        public async Task<PaySpaceApiResponse> SubmitAllCategoriesAsync(int companyId, string frequency, string run, Dictionary<string, List<PayElementEntry>> entriesByCategory)
        {
            var totalResponse = new PaySpaceApiResponse
            {
                Success = true,
                TotalEntries = entriesByCategory.Values.Sum(list => list.Count),
                SuccessfulEntries = 0,
                FailedEntries = 0,
                Errors = new List<string>(),
                Results = new List<PayElementResult>()
            };

            _logger.LogInformation($"Processing all categories for company ID {companyId}");

            // Define the processing order
            var orderedCategories = new[]
            {
        "Hiring",
        "PayElement",
        "Personal Data",
        "Position Data",
        "Pay Rate",
        "Payment Instruction",
        "Employment"
    };

            foreach (var category in orderedCategories)
            {
                if (!entriesByCategory.ContainsKey(category) || !entriesByCategory[category].Any())
                    continue;

                PaySpaceApiResponse response = category switch
                {
                    "Hiring" => await SubmitEmployeeDataAsync(companyId, entriesByCategory[category]),
                    "PayElement" => await SubmitPayElementsAsync(companyId, frequency, run, entriesByCategory[category]),
                    "Personal Data" => await SubmitEmployeeDataAsync(companyId, entriesByCategory[category]),
                    "Position Data" => await SubmitPositionDataAsync(companyId, entriesByCategory[category]),
                    "Pay Rate" => await SubmitPayRateAsync(companyId, entriesByCategory[category]),
                    "Payment Instruction" => await SubmitEmployeeBankingAsync(companyId, entriesByCategory[category]),
                    "Employment" => await SubmitEmploymentStatusAsync(companyId, entriesByCategory[category]),
                    _ => null
                };

                if (response == null)
                    continue;

                MergeResponses(totalResponse, response);

                if (!response.Success)
                {
                    totalResponse.Success = false;
                    totalResponse.Message = $"Failed to process {category}, aborting all other categories";
                    return totalResponse;
                }
            }

            totalResponse.Success = totalResponse.FailedEntries == 0 && totalResponse.SuccessfulEntries > 0;
            totalResponse.Message = totalResponse.Success
                ? $"Successfully processed all {totalResponse.SuccessfulEntries} entries across all categories"
                : $"Processed with {totalResponse.FailedEntries} failures and {totalResponse.SuccessfulEntries} successes across all categories";

            return totalResponse;
        }

        // Helper method to merge category responses
        private void MergeResponses(PaySpaceApiResponse target, PaySpaceApiResponse source)
        {
            target.SuccessfulEntries += source.SuccessfulEntries;
            target.FailedEntries += source.FailedEntries;
            target.Errors.AddRange(source.Errors);
            target.Results.AddRange(source.Results);

            // Only update success if source failed (keep failure state)
            if (!source.Success)
                target.Success = false;
        }

        public async Task<bool> EmployeeExistsAsync(int companyId, string employeeNumber)
        {
            try
            {
                var employee = _payspaceApi.GetEmployeeByEmployeeNumber(companyId, employeeNumber);
                return employee != null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error checking if employee {employeeNumber} exists in company {companyId}");
                return false;
            }
        }

        public async Task<Employee> GetEmployeeAsync(int companyId, string employeeNumber)
        {
            try
            {
                return _payspaceApi.GetEmployeeByEmployeeNumber(companyId, employeeNumber);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting employee {employeeNumber} from company {companyId}");
                return null;
            }
        }

        public async Task<EmployeePosition> GetEmployeePositionAsync(int companyId, string employeeNumber)
        {
            try
            {
                string companyCode = GetCompanyCode(companyId);
                if (string.IsNullOrEmpty(companyCode))
                    return null;

                return _payspaceApi.GetEmployeePositionByEmployeeNumber(employeeNumber, companyCode);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting employee position for {employeeNumber} from company {companyId}");
                return null;
            }
        }

        public async Task<EmployeePayRate> GetEmployeePayRateAsync(int companyId, string employeeNumber)
        {
            try
            {
                string companyCode = GetCompanyCode(companyId);
                if (string.IsNullOrEmpty(companyCode))
                    return null;

                return _payspaceApi.GetEmployeePayrateByEmployeeNumber(employeeNumber, companyCode);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting employee pay rate for {employeeNumber} from company {companyId}");
                return null;
            }
        }

        // Employee Personal Data
        public async Task<PaySpaceApiResponse> SubmitEmployeeDataAsync(int companyId, List<PayElementEntry> entries)
        {
            var response = new PaySpaceApiResponse
            {
                Success = false,
                TotalEntries = entries.Count,
                SuccessfulEntries = 0,
                FailedEntries = 0,
                Errors = new List<string>(),
                Results = new List<PayElementResult>()
            };

            try
            {
                _logger.LogInformation($"Processing employee personal data for company ID {companyId}");

                string companyCode = GetCompanyCode(companyId);
                if (string.IsNullOrEmpty(companyCode))
                {
                    response.Message = "Could not find company code";
                    response.FailedEntries = entries.Count;
                    response.Errors.Add(response.Message);
                    return response;
                }

                var entriesByEmployee = entries.GroupBy(e => e.EmployeeId).ToList();
                _logger.LogInformation($"Processing {entriesByEmployee.Count} employees with personal data changes");

                foreach (var employeeGroup in entriesByEmployee)
                {
                    string employeeId = employeeGroup.Key;
                    var employeeEntries = employeeGroup.ToList();
                    string action = employeeEntries.First().Action;

                    try
                    {
                        var employee = _payspaceApi.GetEmployeeByEmployeeNumber(companyId, employeeId);

                        if (employee == null && action != "ADD")
                        {
                            throw new Exception($"Employee {employeeId} not found and action is not ADD");
                        }

                        if (employee == null)
                        {
                            employee = new Employee
                            {
                                EmployeeNumber = employeeId,
                                DateCreated = DateTime.Today
                            };
                        }

                        // Process all personal data fields
                        foreach (var entry in employeeEntries)
                        {
                            DateTime? effectiveDate = entry.StartDate != default ? entry.StartDate : (DateTime?)null;

                            foreach (var field in entry.CustomFields)
                            {
                                string fieldName = field.Key;
                                string fieldValue = field.Value;

                                switch (fieldName)
                                {
                                    case "Employee Number":
                                        employee.EmployeeNumber = fieldValue;
                                        break;
                                    case "Last Name":
                                        employee.LastName = fieldValue;
                                        break;
                                    case "First Name":
                                        employee.FirstName = fieldValue;
                                        break;
                                    case "Middle Name":
                                        employee.MiddleName = fieldValue;
                                        break;
                                    case "Initials":
                                        employee.Initials = fieldValue;
                                        break;
                                    case "Preferred Name":
                                        employee.PreferredName = fieldValue;
                                        break;
                                    case "Maiden Name":
                                        employee.MaidenName = fieldValue;
                                        break;
                                    case "Title":
                                        employee.Title = fieldValue;
                                        break;
                                    case "Language":
                                        employee.Language = fieldValue;
                                        break;
                                    case "Gender":
                                        if (Enum.TryParse<Gender>(fieldValue, true, out Gender gender))
                                            employee.Gender = gender;
                                        break;
                                    case "Race":
                                        employee.Race = fieldValue;
                                        break;
                                    case "Nationality":
                                        employee.Nationality = fieldValue;
                                        break;
                                    case "Citizenship":
                                        employee.Citizenship = fieldValue;
                                        break;
                                    case "Disability Type":
                                        employee.DisabledType = fieldValue;
                                        break;
                                    case "Marital Status Code":
                                        employee.MaritalStatus = fieldValue;
                                        break;
                                    case "Birth Date":
                                        if (DateTime.TryParse(fieldValue, out DateTime birthDate))
                                            employee.Birthday = birthDate;
                                        break;
                                    case "Home Number":
                                        employee.HomeNumber = fieldValue;
                                        break;
                                    case "Work Number":
                                        employee.WorkNumber = fieldValue;
                                        break;
                                    case "Work Extenstion":
                                        employee.WorkExtension = fieldValue;
                                        break;
                                    case "Cell Number":
                                        employee.CellNumber = fieldValue;
                                        break;
                                    case "Email Address":
                                        employee.Email = fieldValue;
                                        break;
                                    case "Emergency Contact Name":
                                        employee.EmergencyContactName = fieldValue;
                                        break;
                                    case "Eergency Contact Number":
                                        employee.EmergencyContactNumber = fieldValue;
                                        break;
                                    case "Emergency Contact Address":
                                        employee.EmergencyContactAddress = fieldValue;
                                        break;

                                    default:
                                        // Store unknown fields in available custom fields
                                        if (string.IsNullOrEmpty(employee.CustomFieldValue))
                                            employee.CustomFieldValue = $"{fieldName}:{fieldValue}";
                                        else if (string.IsNullOrEmpty(employee.CustomFieldValue2))
                                            employee.CustomFieldValue2 = $"{fieldName}:{fieldValue}";
                                        else
                                            _logger.LogWarning($"No space for custom field {fieldName} for employee {employeeId}");
                                        break;
                                }
                            }
                        }

                        List<string> errorMessages = new List<string>();
                        bool success = _payspaceApi.UpsertEmployeeRecord(employee, ref errorMessages, companyCode);

                        if (success)
                        {
                            response.SuccessfulEntries++;
                            response.Results.Add(new PayElementResult
                            {
                                EmployeeId = employeeId,
                                PayElementId = "PersonalData",
                                Success = true
                            });
                        }
                        else
                        {
                            response.FailedEntries++;
                            response.Errors.AddRange(errorMessages);
                            response.Results.Add(new PayElementResult
                            {
                                EmployeeId = employeeId,
                                PayElementId = "PersonalData",
                                Success = false,
                                ErrorMessage = string.Join("; ", errorMessages)
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        response.FailedEntries++;
                        string errorMessage = $"Error processing personal data for employee {employeeId}: {ex.Message}";
                        _logger.LogError(ex, errorMessage);
                        response.Errors.Add(errorMessage);
                        response.Results.Add(new PayElementResult
                        {
                            EmployeeId = employeeId,
                            PayElementId = "PersonalData",
                            Success = false,
                            ErrorMessage = errorMessage
                        });
                    }
                }

                response.Success = response.FailedEntries == 0 && response.SuccessfulEntries > 0;
                response.Message = response.Success
                    ? $"Successfully processed all {response.SuccessfulEntries} personal data entries"
                    : $"Processed with {response.FailedEntries} failures and {response.SuccessfulEntries} successes";

                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error submitting employee data to PaySpace");
                response.Message = $"Error: {ex.Message}";
                response.FailedEntries = entries.Count - response.SuccessfulEntries;
                response.Errors.Add(ex.Message);
                return response;
            }
        }

        // Payment Instructions - Banking Details
        public async Task<PaySpaceApiResponse> SubmitEmployeeBankingAsync(int companyId, List<PayElementEntry> entries)
        {
            var response = new PaySpaceApiResponse
            {
                Success = false,
                TotalEntries = entries.Count,
                SuccessfulEntries = 0,
                FailedEntries = 0,
                Errors = new List<string>(),
                Results = new List<PayElementResult>()
            };

            try
            {
                _logger.LogInformation($"Processing banking data for company ID {companyId}");

                string companyCode = GetCompanyCode(companyId);
                if (string.IsNullOrEmpty(companyCode))
                {
                    response.Message = "Could not find company code";
                    response.FailedEntries = entries.Count;
                    response.Errors.Add(response.Message);
                    return response;
                }

                var entriesByEmployee = entries.GroupBy(e => e.EmployeeId).ToList();
                _logger.LogInformation($"Processing {entriesByEmployee.Count} employees with banking changes");

                foreach (var employeeGroup in entriesByEmployee)
                {
                    string employeeId = employeeGroup.Key;
                    var employeeEntries = employeeGroup.ToList();
                    string action = employeeEntries.First().Action;

                    try
                    {
                        // Get existing bank details or create new
                        var bankDetail = _payspaceApi.GetEmployeeBankByEmployeeNumber(employeeId, companyCode);

                        if (bankDetail == null && action != "ADD")
                        {
                            throw new Exception($"Bank detail for employee {employeeId} not found and action is not ADD");
                        }

                        if (bankDetail == null)
                        {
                            bankDetail = new EmployeeBankDetail
                            {
                                EmployeeNumber = employeeId
                            };
                        }

                        // Process all banking fields
                        foreach (var entry in employeeEntries)
                        {
                            foreach (var field in entry.CustomFields)
                            {
                                string fieldName = field.Key;
                                string fieldValue = field.Value;

                                // Map all possible payment instruction fields to bank properties
                                switch (fieldName)
                                {
                                    case "Payment Method":
                                        if (fieldValue == "EFT")
                                            bankDetail.PaymentMethod = EmployeePaymentMethod.EFT;
                                        else if (fieldValue == "Cash")
                                            bankDetail.PaymentMethod = EmployeePaymentMethod.Cash;
                                        else if (fieldValue == "Cheque")
                                            bankDetail.PaymentMethod = EmployeePaymentMethod.Cheque;
                                        break;
                                    case "Bank Account Owner":
                                        if (fieldValue == "Own")
                                            bankDetail.BankAccountOwner = BankAccountOwnerType.Own;
                                        else if (fieldValue == "Joint")
                                            bankDetail.BankAccountOwner = BankAccountOwnerType.Joint;
                                        else if (fieldValue == "Third Party")
                                            bankDetail.BankAccountOwner = BankAccountOwnerType.Third_Party;
                                        break;
                                    case "Account Type":
                                        try
                                        {
                                            var accountType = Enum.Parse<AccountType>(fieldValue.Replace(" ", ""), true);
                                            bankDetail.AccountType = accountType;
                                        }
                                        catch
                                        {
                                            _logger.LogWarning($"Could not parse account type: {fieldValue}");
                                        }
                                        break;
                                    case "Bank Name":
                                        bankDetail.BankName = fieldValue;
                                        break;
                                    case "Branch Code":
                                        bankDetail.BankBranchNo = fieldValue;
                                        break;
                                    case "Account No":
                                        bankDetail.BankAccountNo = fieldValue;
                                        break;
                                    case "Name On Account":
                                        bankDetail.BankAccountOwnerName = fieldValue;
                                        break;
                                    case "Comments":
                                        bankDetail.Comments = fieldValue;
                                        break;
                                    case "Swift Code":
                                        bankDetail.SwiftCode = fieldValue;
                                        break;
                                    case "Routing Code":
                                        bankDetail.RoutingCode = fieldValue;
                                        break;


                                    default:
                                        _logger.LogWarning($"Unknown banking field: {fieldName}");
                                        break;
                                }
                            }
                        }

                        // Save changes
                        List<string> errorMessages = new List<string>();
                        bool success = _payspaceApi.UpsertEmployeeBankRecord(bankDetail, ref errorMessages, companyCode);

                        if (success)
                        {
                            response.SuccessfulEntries++;
                            response.Results.Add(new PayElementResult
                            {
                                EmployeeId = employeeId,
                                PayElementId = "BankingDetails",
                                Success = true
                            });
                        }
                        else
                        {
                            response.FailedEntries++;
                            response.Errors.AddRange(errorMessages);
                            response.Results.Add(new PayElementResult
                            {
                                EmployeeId = employeeId,
                                PayElementId = "BankingDetails",
                                Success = false,
                                ErrorMessage = string.Join("; ", errorMessages)
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        response.FailedEntries++;
                        string errorMessage = $"Error processing banking data for employee {employeeId}: {ex.Message}";
                        _logger.LogError(ex, errorMessage);
                        response.Errors.Add(errorMessage);
                        response.Results.Add(new PayElementResult
                        {
                            EmployeeId = employeeId,
                            PayElementId = "BankingDetails",
                            Success = false,
                            ErrorMessage = errorMessage
                        });
                    }
                }

                response.Success = response.FailedEntries == 0 && response.SuccessfulEntries > 0;
                response.Message = response.Success
                    ? $"Successfully processed all {response.SuccessfulEntries} banking entries"
                    : $"Processed with {response.FailedEntries} failures and {response.SuccessfulEntries} successes";

                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error submitting banking data to PaySpace");
                response.Message = $"Error: {ex.Message}";
                response.FailedEntries = entries.Count - response.SuccessfulEntries;
                response.Errors.Add(ex.Message);
                return response;
            }
        }

        // Employment Status
        public async Task<PaySpaceApiResponse> SubmitEmploymentStatusAsync(int companyId, List<PayElementEntry> entries)
        {
            var response = new PaySpaceApiResponse
            {
                Success = false,
                TotalEntries = entries.Count,
                SuccessfulEntries = 0,
                FailedEntries = 0,
                Errors = new List<string>(),
                Results = new List<PayElementResult>()
            };

            try
            {
                _logger.LogInformation($"Processing employment status data for company ID {companyId}");

                string companyCode = GetCompanyCode(companyId);
                if (string.IsNullOrEmpty(companyCode))
                {
                    response.Message = "Could not find company code";
                    response.FailedEntries = entries.Count;
                    response.Errors.Add(response.Message);
                    return response;
                }

                var entriesByEmployee = entries.GroupBy(e => e.EmployeeId).ToList();
                _logger.LogInformation($"Processing {entriesByEmployee.Count} employees with employment status changes");

                foreach (var employeeGroup in entriesByEmployee)
                {
                    string employeeId = employeeGroup.Key;
                    var employeeEntries = employeeGroup.ToList();
                    string action = employeeEntries.First().Action;

                    try
                    {
                        // Get existing employment status or create new
                        var employmentStatus = _payspaceApi.GetEmployeeEmploymentStatusseByEmployeeId(employeeId, companyCode);

                        if (employmentStatus == null && action != "ADD")
                        {
                            throw new Exception($"Employment status for employee {employeeId} not found and action is not ADD");
                        }

                        if (employmentStatus == null)
                        {
                            employmentStatus = new EmployeeEmploymentStatus
                            {
                                EmployeeNumber = employeeId,
                                EmploymentDate = DateTime.Today
                            };
                        }

                        // Set appropriate employment action based on the action type
                        if (action == "ADD" || action == "CHANGE")
                        {
                            employmentStatus.EmploymentAction = EmploymentAction.New;
                        }
                        else if (action == "TERMINATE")
                        {
                            employmentStatus.EmploymentAction = EmploymentAction.Terminatethisemployee;

                            if (!employmentStatus.TerminationDate.HasValue)
                            {
                                // Default to today if not set by a field
                                employmentStatus.TerminationDate = DateTime.Today;
                            }
                        }
                        else if (action.Contains("REINSTATE"))
                        {
                            // reinstatement cases
                            if (action.Contains("NEWTAX"))
                            {
                                employmentStatus.EmploymentAction = EmploymentAction.Reinstatethisemployeestartinganewtaxrecord;
                            }
                            else
                            {
                                employmentStatus.EmploymentAction = EmploymentAction.Reinstatethisemployeeresumingthistaxrecord;
                            }
                        }

                        // Process all employment status fields
                        foreach (var entry in employeeEntries)
                        {
                            if (entry.StartDate != default)
                            {
                                employmentStatus.EmploymentDate = entry.StartDate;
                            }

                            // Process each custom field
                            foreach (var field in entry.CustomFields)
                            {
                                string fieldName = field.Key;
                                string fieldValue = field.Value;

                                switch (fieldName)
                                {
                                    case "Group Join Date":
                                        if (DateTime.TryParse(fieldValue, out DateTime groupJoinDate))
                                            employmentStatus.GroupJoinDate = groupJoinDate;
                                        break;
                                    case "Start Date":
                                        if (DateTime.TryParse(fieldValue, out DateTime startDate))
                                            employmentStatus.EmploymentDate = startDate;
                                        break;
                                    case "Termination Date":
                                        if (DateTime.TryParse(fieldValue, out DateTime terminationDate))
                                            employmentStatus.TerminationDate = terminationDate;
                                        break;
                                    case "Termination Reason":
                                        employmentStatus.TerminationReason = fieldValue;
                                        break;
                                    case "Nature Of Person":
                                        employmentStatus.NatureOfPerson = fieldValue;
                                        break;
                                    case "Identity Type":
                                        employmentStatus.IdentityType = fieldValue;
                                        break;
                                    case "ID Number":
                                        employmentStatus.IdNumber = fieldValue;
                                        break;
                                    case "Passport Issuing Country":
                                        employmentStatus.PassportCountry = fieldValue;
                                        break;
                                    case "Passport Number":
                                        employmentStatus.PassportNumber = fieldValue;
                                        break;
                                    case "Passport Issue Date":
                                        if (DateTime.TryParse(fieldValue, out DateTime passportIssueDate))
                                            employmentStatus.PassportIssued = passportIssueDate;
                                        break;
                                    case "Passport Expiry Date":
                                        if (DateTime.TryParse(fieldValue, out DateTime passportExpiryDate))
                                            employmentStatus.PassportExpiry = passportExpiryDate;
                                        break;
                                    case "Tax Status":
                                        employmentStatus.TaxStatus = fieldValue;
                                        break;
                                    case "Tax Ref. Number":
                                        employmentStatus.TaxReferenceNumber = fieldValue;
                                        break;
                                    case "Reference Number":
                                        employmentStatus.ReferenceNumber = fieldValue;
                                        break;
                                    default:
                                        _logger.LogWarning($"Unknown employment field: {fieldName}");
                                        break;
                                }
                            }
                        }

                        // Save changes
                        List<string> errorMessages = new List<string>();
                        bool success = _payspaceApi.UpsertEmployeeEmploymentStatusRecord(employmentStatus, ref errorMessages, companyCode);

                        if (success)
                        {
                            response.SuccessfulEntries++;
                            response.Results.Add(new PayElementResult
                            {
                                EmployeeId = employeeId,
                                PayElementId = "EmploymentStatus",
                                Success = true
                            });
                        }
                        else
                        {
                            response.FailedEntries++;
                            response.Errors.AddRange(errorMessages);
                            response.Results.Add(new PayElementResult
                            {
                                EmployeeId = employeeId,
                                PayElementId = "EmploymentStatus",
                                Success = false,
                                ErrorMessage = string.Join("; ", errorMessages)
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        response.FailedEntries++;
                        string errorMessage = $"Error processing employment status for employee {employeeId}: {ex.Message}";
                        _logger.LogError(ex, errorMessage);
                        response.Errors.Add(errorMessage);
                        response.Results.Add(new PayElementResult
                        {
                            EmployeeId = employeeId,
                            PayElementId = "EmploymentStatus",
                            Success = false,
                            ErrorMessage = errorMessage
                        });
                    }
                }

                response.Success = response.FailedEntries == 0 && response.SuccessfulEntries > 0;
                response.Message = response.Success
                    ? $"Successfully processed all {response.SuccessfulEntries} employment status entries"
                    : $"Processed with {response.FailedEntries} failures and {response.SuccessfulEntries} successes";

                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error submitting employment status data to PaySpace");
                response.Message = $"Error: {ex.Message}";
                response.FailedEntries = entries.Count - response.SuccessfulEntries;
                response.Errors.Add(ex.Message);
                return response;
            }
        }

        public async Task<PaySpaceApiResponse> SubmitPayRateAsync(int companyId, List<PayElementEntry> entries)
        {
            var response = new PaySpaceApiResponse
            {
                Success = false,
                TotalEntries = entries.Count,
                SuccessfulEntries = 0,
                FailedEntries = 0,
                Errors = new List<string>(),
                Results = new List<PayElementResult>()
            };

            try
            {
                _logger.LogInformation($"Processing pay rate data for company ID {companyId}");

                string companyCode = GetCompanyCode(companyId);
                if (string.IsNullOrEmpty(companyCode))
                {
                    response.Message = "Could not find company code";
                    response.FailedEntries = entries.Count;
                    response.Errors.Add(response.Message);
                    return response;
                }

                var entriesByEmployee = entries.GroupBy(e => e.EmployeeId).ToList();
                _logger.LogInformation($"Processing {entriesByEmployee.Count} employees with pay rate changes");

                foreach (var employeeGroup in entriesByEmployee)
                {
                    string employeeId = employeeGroup.Key;
                    var employeeEntries = employeeGroup.ToList();
                    string action = employeeEntries.First().Action;

                    try
                    {
                        // Get existing pay rate or create new
                        var payRate = _payspaceApi.GetEmployeePayrateByEmployeeNumber(employeeId, companyCode);

                        if (payRate == null && action != "ADD")
                        {
                            throw new Exception($"Pay rate for employee {employeeId} not found and action is not ADD");
                        }

                        if (payRate == null)
                        {
                            payRate = new EmployeePayRate
                            {
                                EmployeeNumber = employeeId,
                                EffectiveDate = DateTime.Today
                            };
                        }

                        // Process all pay rate fields
                        foreach (var entry in employeeEntries)
                        {
                            // Set effective date from entry
                            if (entry.StartDate != default)
                            {
                                payRate.EffectiveDate = entry.StartDate;
                            }

                            foreach (var field in entry.CustomFields)
                            {
                                string fieldName = field.Key;
                                string fieldValue = field.Value;

                                switch (fieldName)
                                {
                                    case "Basic Salary":
                                    case "Salary":
                                        if (decimal.TryParse(fieldValue, out decimal salary))
                                            payRate.Package = salary;
                                        break;
                                    case "Pay Frequency":
                                        payRate.PayFrequency = fieldValue;
                                        break;

                                    case "Automatic Pay Indicator":
                                        if (bool.TryParse(fieldValue, out bool autoPayInd))
                                            payRate.AutomaticPayInd = autoPayInd;
                                        break;
                                    case "Reason":
                                        payRate.Reason = fieldValue;
                                        break;
                                    case "Comments":
                                        payRate.Comments = fieldValue;
                                        break;
                                    case "Effective Date":
                                        if (DateTime.TryParse(fieldValue, out DateTime effDate))
                                            payRate.EffectiveDate = effDate;
                                        break;
                                    default:
                                        _logger.LogWarning($"Unknown pay rate field: {fieldName}");
                                        break;
                                }
                            }
                        }

                        // Save changes
                        List<string> errorMessages = new List<string>();
                        bool success = _payspaceApi.UpsertEmployeePayrateRecord(payRate, ref errorMessages, companyCode);

                        if (success)
                        {
                            response.SuccessfulEntries++;
                            response.Results.Add(new PayElementResult
                            {
                                EmployeeId = employeeId,
                                PayElementId = "PayRate",
                                Success = true
                            });
                        }
                        else
                        {
                            response.FailedEntries++;
                            response.Errors.AddRange(errorMessages);
                            response.Results.Add(new PayElementResult
                            {
                                EmployeeId = employeeId,
                                PayElementId = "PayRate",
                                Success = false,
                                ErrorMessage = string.Join("; ", errorMessages)
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        response.FailedEntries++;
                        string errorMessage = $"Error processing pay rate for employee {employeeId}: {ex.Message}";
                        _logger.LogError(ex, errorMessage);
                        response.Errors.Add(errorMessage);
                        response.Results.Add(new PayElementResult
                        {
                            EmployeeId = employeeId,
                            PayElementId = "PayRate",
                            Success = false,
                            ErrorMessage = errorMessage
                        });
                    }
                }

                response.Success = response.FailedEntries == 0 && response.SuccessfulEntries > 0;
                response.Message = response.Success
                    ? $"Successfully processed all {response.SuccessfulEntries} pay rate entries"
                    : $"Processed with {response.FailedEntries} failures and {response.SuccessfulEntries} successes";

                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error submitting pay rate data to PaySpace");
                response.Message = $"Error: {ex.Message}";
                response.FailedEntries = entries.Count - response.SuccessfulEntries;
                response.Errors.Add(ex.Message);
                return response;
            }
        }

        // Position Data (Deployment and Approver)
        public async Task<PaySpaceApiResponse> SubmitPositionDataAsync(int companyId, List<PayElementEntry> entries)
        {
            var response = new PaySpaceApiResponse
            {
                Success = false,
                TotalEntries = entries.Count,
                SuccessfulEntries = 0,
                FailedEntries = 0,
                Errors = new List<string>(),
                Results = new List<PayElementResult>()
            };

            try
            {
                _logger.LogInformation($"Processing position data for company ID {companyId}");

                string companyCode = GetCompanyCode(companyId);
                if (string.IsNullOrEmpty(companyCode))
                {
                    response.Message = "Could not find company code";
                    response.FailedEntries = entries.Count;
                    response.Errors.Add(response.Message);
                    return response;
                }

                // Group entries by employee and category (to handle Deployment and Approver separately)
                var entriesByEmployeeAndCategory = entries
                    .GroupBy(e => new { e.EmployeeId, e.Category })
                    .ToList();

                _logger.LogInformation($"Processing {entriesByEmployeeAndCategory.Count} employee/category combinations");

                foreach (var group in entriesByEmployeeAndCategory)
                {
                    string employeeId = group.Key.EmployeeId;
                    string category = group.Key.Category;
                    var categoryEntries = group.ToList();
                    string action = categoryEntries.First().Action;

                    try
                    {
                        // Get existing position or create new
                        var position = _payspaceApi.GetEmployeePositionByEmployeeNumber(employeeId, companyCode);

                        if (position == null && action != "ADD")
                        {
                            throw new Exception($"Position for employee {employeeId} not found and action is not ADD");
                        }

                        if (position == null)
                        {
                            position = new EmployeePosition
                            {
                                EmployeeNumber = employeeId,
                                EffectiveDate = DateTime.Today
                            };
                        }

                        // Update effective date from entries if available
                        foreach (var entry in categoryEntries)
                        {
                            if (entry.StartDate != default)
                            {
                                position.EffectiveDate = entry.StartDate;
                                break;
                            }
                        }

                        // Process position fields based on category
                        foreach (var entry in categoryEntries)
                        {
                            foreach (var field in entry.CustomFields)
                            {
                                string fieldName = field.Key;
                                string fieldValue = field.Value;

                                if (category == "Deployment")
                                {
                                    // Map Deployment fields
                                    switch (fieldName)
                                    {
                                        case "Position Title":
                                            position.OrganizationPosition = fieldValue;
                                            break;
                                        case "Position Type":
                                            position.PositionType = fieldValue;
                                            break;
                                        case "Grade":
                                            position.Grade = fieldValue;
                                            break;
                                        case "Organization Group":
                                            position.OrganizationGroup = fieldValue;
                                            break;
                                        case "Organization Region":
                                            position.OrganizationRegion = fieldValue;
                                            break;
                                        case "Position Effective Date":
                                            if (DateTime.TryParse(fieldValue, out DateTime posEffDate))
                                                position.PositionEffectiveDate = posEffDate;
                                            break;
                                        case "Employment Category":
                                            position.EmploymentCategory = fieldValue;
                                            break;
                                        case "Employment Sub Category":
                                            position.EmploymentSubCategory = fieldValue;
                                            break;
                                        case "Job":
                                            position.Job = fieldValue;
                                            break;
                                        case "AltPositionName":
                                            position.AltPositionName = fieldValue;
                                            break;
                                        case "Roster":
                                            position.Roster = fieldValue;
                                            break;
                                        case "TradeUnion":
                                            position.TradeUnion = fieldValue;
                                            break;
                                        case "WorkflowRole":
                                            position.WorkflowRole = fieldValue;
                                            break;
                                        case "Comments":
                                            position.Comments = fieldValue;
                                            break;
                                        default:
                                            _logger.LogWarning($"Unknown deployment field: {fieldName}");
                                            break;
                                    }
                                }
                                else if (category == "Approver")
                                {
                                    // Map Approver fields
                                    switch (fieldName)
                                    {
                                        case "HiIs":
                                            position.DirectlyReportsEmployeeNumber = fieldValue;
                                            break;
                                        case "Approver Type":
                                            position.DirectlyReportsPosition = fieldValue;
                                            break;
                                        case "Reports To":
                                            position.DirectlyReportsEmployee = fieldValue;
                                            break;
                                        case "DirectlyReportsPositionOverride":
                                            position.DirectlyReportsPositionOverride = fieldValue;
                                            break;
                                        case "Administrator":
                                            position.Administrator = fieldValue;
                                            break;
                                        case "AdministratorEmployeeNumber":
                                            position.AdministratorEmployeeNumber = fieldValue;
                                            break;
                                        default:
                                            _logger.LogWarning($"Unknown approver field: {fieldName}");
                                            break;
                                    }
                                }
                            }
                        }

                        // Save changes
                        List<string> errorMessages = new List<string>();
                        bool success = _payspaceApi.UpsertEmployeePositionRecord(position, ref errorMessages, companyCode);

                        if (success)
                        {
                            response.SuccessfulEntries += categoryEntries.Count;
                            response.Results.Add(new PayElementResult
                            {
                                EmployeeId = employeeId,
                                PayElementId = category,
                                ComponentCode = position.OrganizationPosition ?? "",
                                Success = true
                            });
                        }
                        else
                        {
                            response.FailedEntries += categoryEntries.Count;
                            response.Errors.AddRange(errorMessages);
                            response.Results.Add(new PayElementResult
                            {
                                EmployeeId = employeeId,
                                PayElementId = category,
                                ComponentCode = position.OrganizationPosition ?? "",
                                Success = false,
                                ErrorMessage = string.Join("; ", errorMessages)
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        response.FailedEntries += categoryEntries.Count;
                        string errorMessage = $"Error processing {category} for employee {employeeId}: {ex.Message}";
                        _logger.LogError(ex, errorMessage);
                        response.Errors.Add(errorMessage);
                        response.Results.Add(new PayElementResult
                        {
                            EmployeeId = employeeId,
                            PayElementId = category,
                            Success = false,
                            ErrorMessage = errorMessage
                        });
                    }
                }

                response.Success = response.FailedEntries == 0 && response.SuccessfulEntries > 0;
                response.Message = response.Success
                    ? $"Successfully processed all {response.SuccessfulEntries} position entries"
                    : $"Processed with {response.FailedEntries} failures and {response.SuccessfulEntries} successes";

                return response;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error submitting position data to PaySpace");
                response.Message = $"Error: {ex.Message}";
                response.FailedEntries = entries.Count - response.SuccessfulEntries;
                response.Errors.Add(ex.Message);
                return response;
            }
        }

        private string GetCompanyCode(int companyId)
        {
            try
            {
                var agencyCompanies = _payspaceApi.GetAgencyCompanies();
                var company = agencyCompanies
                    .SelectMany(a => a.companies)
                    .FirstOrDefault(c => c.company_id == companyId);

                return company?.company_code;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error getting company code for ID {companyId}");
                return null;
            }
        }
    }
}
