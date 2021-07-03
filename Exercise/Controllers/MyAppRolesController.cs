using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Exercise.Controllers
{
    public class MyAppRolesController : Controller
    {
        public UserManager<IdentityUser> UserManager;
        public RoleManager<IdentityRole> RoleManager;


        public MyAppRolesController(RoleManager<IdentityRole> _RoleManager,
          UserManager<IdentityUser> _userManager)
        {
            UserManager = _userManager;
            RoleManager = _RoleManager;

        }

        public async Task<IActionResult> Index()
        {
            IdentityRole Adminrole = new IdentityRole { Name = "Admin" };
            IdentityRole Guestrole = new IdentityRole { Name = "Guest" };

            IdentityResult result = await RoleManager.CreateAsync(Adminrole);
            IdentityResult result2 = await RoleManager.CreateAsync(Guestrole);

           // IdentityUser user1 = await UserManager.FindByIdAsync("2e78f36e-5cd0-4db4-8eb2-8eba03e1c584"); //nabeel
           // result = await UserManager.AddToRoleAsync(user1, "Admin");

            IdentityUser user2 = await UserManager.FindByIdAsync("3e0d2757-71ec-49c7-a9a3-95ac12291048"); //zarah
            result = await UserManager.AddToRoleAsync(user2, "Guest");


            if (result.Succeeded)
            {
                return RedirectToPage("/");
            }
            return View();
        }
    }
}
