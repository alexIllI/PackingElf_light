# root/tests/test_account_manager.py
import pytest
import os
import sys

# Add the root directory to the sys.path to enable module imports from the root level
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from operation.account_manage import EncryptedAccountManager

@pytest.fixture
def account_manager():
    manager = EncryptedAccountManager()
    yield manager
    if os.path.exists(manager.file_path):
        os.remove(manager.file_path)

def test_add_and_get_account(account_manager):
    username = "test_user"
    account_info = {"email": "test_user@example.com", "password": "password123"}
    account_manager.add_account(username, account_info)
    
    assert account_manager.get_account_by_name(username) == account_info

def test_encrypt_and_save(account_manager):
    username = "test_user"
    account_info = {"email": "test_user@example.com", "password": "password123"}
    account_manager.add_account(username, account_info)
    account_manager.encrypt_and_save()
    
    assert os.path.exists(account_manager.file_path)

def test_load_and_decrypt(account_manager):
    username = "test_user"
    account_info = {"email": "test_user@example.com", "password": "password123"}
    account_manager.add_account(username, account_info)
    account_manager.encrypt_and_save()
    
    new_manager = EncryptedAccountManager()
    new_manager.load_and_decrypt()
    
    assert new_manager.get_account_by_name(username) == account_info

def test_update_account_by_name(account_manager):
    username = "test_user"
    account_info = {"email": "test_user@example.com", "password": "password123"}
    updated_account_info = {"email": "updated_user@example.com", "password": "newpassword456"}
    account_manager.add_account(username, account_info)
    account_manager.update_account_by_name(username, updated_account_info)
    
    assert account_manager.get_account_by_name(username) == updated_account_info

def test_delete_account_by_name(account_manager):
    username = "test_user"
    account_info = {"email": "test_user@example.com", "password": "password123"}
    account_manager.add_account(username, account_info)
    account_manager.delete_account_by_name(username)
    
    assert account_manager.get_account_by_name(username) is None

def test_get_all_account_names(account_manager):
    accounts = {
        "user1": {"email": "user1@example.com", "password": "password123"},
        "user2": {"email": "user2@example.com", "password": "password456"}
    }
    for username, account_info in accounts.items():
        account_manager.add_account(username, account_info)
    
    assert set(account_manager.get_all_account_names()) == set(accounts.keys())

if __name__ == "__main__":
    pytest.main()
