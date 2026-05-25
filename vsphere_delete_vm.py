#!/usr/bin/env python3
"""
vSphere VM Delete Tool (Interactive)
Connects to vCenter and allows you to search and delete VMs interactively.
- Powers off the VM if it's running
- Deletes the VM and removes all associated disk files
"""

import requests
import urllib3
import time
import sys

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ─── Configuration ───────────────────────────────────────────────────────────
VCENTER = "vcenter.ce.te-labs.1keyes.net"
USERNAME = "moenassa"
PASSWORD = "C!$c0_123456"
# ─────────────────────────────────────────────────────────────────────────────

# Dry-run mode: pass --dry-run as argument to simulate without making changes
DRY_RUN = "--dry-run" in sys.argv

BASE_URL = f"https://{VCENTER}/api"
session = requests.Session()
session.verify = False


def authenticate():
    """Authenticate and get session token."""
    resp = session.post(f"{BASE_URL}/session", auth=(USERNAME, PASSWORD))
    if resp.status_code != 201:
        print(f"[ERROR] Authentication failed: {resp.status_code} - {resp.text}")
        exit(1)
    token = resp.json()
    session.headers["vmware-api-session-id"] = token
    print(f"[+] Authenticated to {VCENTER}\n")


def logout():
    """End the session."""
    session.delete(f"{BASE_URL}/session")


def search_vms(search_term):
    """Search for VMs matching the given name (case-insensitive partial match)."""
    resp = session.get(f"{BASE_URL}/vcenter/vm")
    if resp.status_code != 200:
        print(f"[ERROR] Failed to list VMs: {resp.status_code}")
        return []

    all_vms = resp.json()
    matches = [vm for vm in all_vms if search_term.lower() in vm['name'].lower()]
    return matches


def get_vm_details(vm_id):
    """Get VM details."""
    resp = session.get(f"{BASE_URL}/vcenter/vm/{vm_id}")
    if resp.status_code != 200:
        return None
    return resp.json()


def power_off_vm(vm_id, vm_name):
    """Power off a VM (hard stop)."""
    if DRY_RUN:
        print(f"    [DRY-RUN] Would power off '{vm_name}' (no action taken)")
        return True

    print(f"    [*] Powering off '{vm_name}'...")

    # Try guest shutdown first
    resp = session.post(f"{BASE_URL}/vcenter/vm/{vm_id}/guest/power?action=shutdown")
    if resp.status_code == 204:
        print(f"    [*] Guest shutdown initiated, waiting up to 60s...")
        for i in range(12):
            time.sleep(5)
            detail = get_vm_details(vm_id)
            if detail and detail.get('power_state') == 'POWERED_OFF':
                print(f"    [+] VM gracefully shut down.")
                return True
        print(f"    [!] Guest shutdown timed out, forcing power off...")

    # Force power off
    resp = session.post(f"{BASE_URL}/vcenter/vm/{vm_id}/power?action=stop")
    if resp.status_code == 204:
        print(f"    [+] VM powered off (hard stop).")
        time.sleep(2)
        return True
    elif resp.status_code == 400:
        # Already powered off
        return True
    else:
        print(f"    [ERROR] Failed to power off: {resp.status_code} - {resp.text}")
        return False


def delete_vm(vm_id, vm_name):
    """Delete a VM (removes from disk)."""
    if DRY_RUN:
        print(f"    [DRY-RUN] Would delete VM '{vm_name}' (ID: {vm_id}) and all disk files (no action taken)")
        return True

    print(f"    [*] Deleting VM '{vm_name}' and all associated files...")
    resp = session.delete(f"{BASE_URL}/vcenter/vm/{vm_id}")
    if resp.status_code == 204:
        print(f"    [+] VM '{vm_name}' deleted successfully.")
        return True
    else:
        print(f"    [ERROR] Failed to delete: {resp.status_code} - {resp.text}")
        return False


def display_vm_info(vm):
    """Display VM info in a formatted way."""
    state_icon = "ON" if vm['power_state'] == "POWERED_ON" else "OFF"
    print(f"    Name:   {vm['name']}")
    print(f"    ID:     {vm['vm']}")
    print(f"    State:  {state_icon}")
    print(f"    CPU:    {vm.get('cpu_count', '?')} vCPUs")
    print(f"    RAM:    {vm.get('memory_size_MiB', '?')} MB")


def interactive_delete():
    """Main interactive loop for deleting VMs."""
    print("=" * 60)
    print(" vSphere VM Delete Tool")
    if DRY_RUN:
        print(" *** DRY-RUN MODE — no changes will be made ***")
    print(" Type 'quit' or 'exit' to stop")
    print(" Type 'list' to show all VMs")
    print("=" * 60)

    while True:
        print()
        search = input("Enter VM name to delete (or search term): ").strip()

        if not search:
            continue
        if search.lower() in ('quit', 'exit', 'q'):
            print("\n[+] Goodbye.")
            break
        if search.lower() == 'list':
            resp = session.get(f"{BASE_URL}/vcenter/vm")
            if resp.status_code == 200:
                vms = sorted(resp.json(), key=lambda x: x['name'].lower())
                print(f"\n{'#':<4} {'VM Name':<40} {'State':<8} {'CPU':<5} {'RAM MB'}")
                print(f"{'-'*4} {'-'*40} {'-'*8} {'-'*5} {'-'*8}")
                for i, vm in enumerate(vms, 1):
                    state = "ON" if vm['power_state'] == "POWERED_ON" else "OFF"
                    print(f"{i:<4} {vm['name']:<40} {state:<8} {vm.get('cpu_count','?'):<5} {vm.get('memory_size_MiB','?')}")
            continue

        # Search for matching VMs
        matches = search_vms(search)

        if not matches:
            print(f"[!] No VMs found matching '{search}'")
            continue

        if len(matches) == 1:
            vm = matches[0]
            print(f"\n  Found 1 VM:")
            display_vm_info(vm)
        else:
            print(f"\n  Found {len(matches)} VMs matching '{search}':\n")
            for i, vm in enumerate(matches, 1):
                state = "ON" if vm['power_state'] == "POWERED_ON" else "OFF"
                print(f"    [{i}] {vm['name']:<40} ({state}, {vm.get('cpu_count','?')} vCPU, {vm.get('memory_size_MiB','?')} MB)")

            print()
            choice = input("  Select VM number to delete (or 'all' for all, 'cancel' to cancel): ").strip()

            if choice.lower() in ('cancel', 'c', ''):
                print("  Cancelled.")
                continue
            elif choice.lower() == 'all':
                confirm = input(f"\n  *** DELETE ALL {len(matches)} VMs? Type 'YES DELETE ALL' to confirm: ").strip()
                if confirm != "YES DELETE ALL":
                    print("  Cancelled.")
                    continue
                for vm in matches:
                    print(f"\n  --- Deleting: {vm['name']} ---")
                    if vm['power_state'] == "POWERED_ON":
                        if not power_off_vm(vm['vm'], vm['name']):
                            print(f"    [!] Skipping (could not power off)")
                            continue
                    delete_vm(vm['vm'], vm['name'])
                continue
            else:
                try:
                    idx = int(choice) - 1
                    if 0 <= idx < len(matches):
                        vm = matches[idx]
                    else:
                        print("  Invalid selection.")
                        continue
                except ValueError:
                    print("  Invalid input.")
                    continue

        # Confirm deletion
        print()
        confirm = input(f"  *** DELETE '{vm['name']}'? This will remove all disk files. (yes/no): ").strip()

        if confirm.lower() not in ('yes', 'y'):
            print("  Cancelled.")
            continue

        # Power off if running
        if vm['power_state'] == "POWERED_ON":
            if not power_off_vm(vm['vm'], vm['name']):
                print("  [!] Could not power off VM. Aborting delete.")
                continue

        # Delete
        delete_vm(vm['vm'], vm['name'])


def main():
    authenticate()
    try:
        interactive_delete()
    finally:
        logout()
        print("[+] Session closed.")


if __name__ == "__main__":
    main()
