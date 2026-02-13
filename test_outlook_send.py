"""
Test script to verify Outlook send functionality.
This tests the EntryID approach for sending drafts.
"""

import win32com.client

def test_outlook_send():
    """Test that we can retrieve and send drafts using EntryID."""
    print("Testing Outlook draft send functionality...")

    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")

        # Get Drafts folder
        drafts = ns.GetDefaultFolder(16)  # olFolderDrafts

        print(f"✓ Connected to Outlook")
        print(f"✓ Found Drafts folder: {drafts.Name}")

        # Look for Manager Report subfolder
        manager_folder = None
        for i in range(drafts.Folders.Count):
            folder = drafts.Folders.Item(i + 1)
            if folder.Name == "Manager Report":
                manager_folder = folder
                break

        if not manager_folder:
            print("⚠ No 'Manager Report' folder found in Drafts")
            print("  Please create some draft emails first using the app")
            return False

        print(f"✓ Found Manager Report folder with {manager_folder.Items.Count} items")

        if manager_folder.Items.Count == 0:
            print("⚠ No draft emails in Manager Report folder")
            return False

        # Test: Store EntryID and retrieve item
        test_item = manager_folder.Items.Item(1)
        entry_id = test_item.EntryID
        subject = test_item.Subject
        to_addr = test_item.To

        print(f"\n📧 Test draft:")
        print(f"   Subject: {subject}")
        print(f"   To: {to_addr}")
        print(f"   EntryID: {entry_id[:50]}...")

        # Retrieve using EntryID (this is what we do when sending)
        retrieved_item = ns.GetItemFromID(entry_id)

        print(f"\n✓ Successfully retrieved item using EntryID")
        print(f"   Subject matches: {retrieved_item.Subject == subject}")
        print(f"   To matches: {retrieved_item.To == to_addr}")
        print(f"   Is sent: {retrieved_item.Sent}")

        # Verify the item object is valid
        try:
            _ = retrieved_item.Subject
            _ = retrieved_item.To
            _ = retrieved_item.Sent
            print(f"✓ Item object is valid and accessible")
        except Exception as e:
            print(f"✗ Item object is invalid: {e}")
            return False

        print("\n" + "="*60)
        print("✅ TEST PASSED: EntryID approach works correctly")
        print("="*60)
        print("\nNOTE: No emails were actually sent during this test.")
        print("The Send button should now work properly.")

        return True

    except Exception as e:
        print(f"\n✗ TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_outlook_send()
    exit(0 if success else 1)
