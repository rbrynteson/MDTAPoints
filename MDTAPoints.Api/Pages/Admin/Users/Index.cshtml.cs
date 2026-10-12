using System.Security.Claims;
using MDTAPoints.Api.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;

namespace MDTAPoints.Api.Pages.Admin.Users;

public class IndexModel : PageModel
{
    private readonly CupDatabase _database;

    public IndexModel(CupDatabase database)
    {
        _database = database;
    }

    [BindProperty]
    public string Username { get; set; } = "";

    [BindProperty]
    public string DisplayName { get; set; } = "";

    [BindProperty]
    public string Password { get; set; } = "";

    public IReadOnlyList<AppUser> Users { get; private set; } = Array.Empty<AppUser>();
    public string? Error { get; private set; }

    public async Task OnGetAsync()
    {
        Users = await _database.ListUsersAsync(HttpContext.RequestAborted);
    }

    public async Task<IActionResult> OnPostAsync()
    {
        try
        {
            await _database.AddUserAsync(Username, DisplayName, Password, HttpContext.RequestAborted);
        }
        catch (SqlException)
        {
            Error = "That username is already in use.";
            Users = await _database.ListUsersAsync(HttpContext.RequestAborted);
            return Page();
        }

        return RedirectToPage();
    }

    public async Task<IActionResult> OnPostDeleteAsync(int id)
    {
        var current = int.Parse(User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "0");
        var removed = await _database.DeleteUserAsync(id, current, HttpContext.RequestAborted);
        if (!removed)
        {
            Error = "You cannot remove the account you are using.";
            Users = await _database.ListUsersAsync(HttpContext.RequestAborted);
            return Page();
        }

        return RedirectToPage();
    }
}
