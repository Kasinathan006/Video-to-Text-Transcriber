#!/usr/bin/env python3
"""
manage.py - VoxDoc AI seller & operations CLI

Commands:
    python manage.py init                                  Initialize the database
    python manage.py create-admin <email> <password>       Create an admin account
    python manage.py gen-keys <pro|agency> [--days 30] [--count 5]
                                                           Mint sellable license keys
    python manage.py list-keys [--unredeemed]              Show license keys
    python manage.py list-users                            Show all accounts & usage
    python manage.py set-tier <email> <free|pro|agency> [--days 30]
                                                           Manually change a user's plan
"""

import argparse
import sys
import time

import database as db


def cmd_init(_args):
    db.init_db()
    print(f"Database ready at: {db.DB_PATH}")


def cmd_create_admin(args):
    db.init_db()
    user = db.create_user(args.email, args.password, full_name="Administrator",
                          tier="agency", is_admin=True)
    db.set_user_tier(user["id"], "agency", duration_days=None)  # never expires
    print(f"Admin account created: {user['email']} (tier: agency, never expires)")
    print("Log in at http://localhost:8000/app to access the admin panel.")


def cmd_gen_keys(args):
    db.init_db()
    keys = db.generate_license_keys(args.tier, duration_days=args.days, count=args.count)
    print(f"Minted {len(keys)} '{args.tier}' license key(s), each valid {args.days} days "
          f"from redemption:\n")
    for k in keys:
        print(f"  {k}")
    print("\nSell these via Gumroad / Lemon Squeezy / invoice. "
          "Buyers redeem them on their Account page.")


def cmd_list_keys(args):
    db.init_db()
    keys = db.list_license_keys(include_redeemed=not args.unredeemed)
    if not keys:
        print("No license keys found.")
        return
    for k in keys:
        state = "REDEEMED" if k["redeemed_by"] else "available"
        print(f"  {k['key']}  tier={k['tier']:<7} {k['duration_days']}d  [{state}]")


def cmd_list_users(_args):
    db.init_db()
    users = db.list_users()
    if not users:
        print("No users yet.")
        return
    print(f"{'EMAIL':<36} {'TIER':<8} {'ADMIN':<6} {'USED (min)':<11} EXPIRES")
    for u in users:
        q = db.quota_summary(u)
        exp = "-"
        if u["tier_expires_at"]:
            exp = time.strftime("%Y-%m-%d", time.localtime(u["tier_expires_at"]))
        print(f"{u['email']:<36} {q['tier']:<8} {'yes' if u['is_admin'] else 'no':<6} "
              f"{q['minutes_used']:<11} {exp}")


def cmd_set_tier(args):
    db.init_db()
    target = next((u for u in db.list_users() if u["email"] == args.email.lower().strip()), None)
    if not target:
        print(f"Error: no user with email '{args.email}'.")
        sys.exit(1)
    days = None if args.days == 0 else args.days
    db.set_user_tier(target["id"], args.tier, duration_days=days)
    life = "never expires" if days is None else f"{days} days"
    print(f"{args.email} is now on '{args.tier}' ({life}).")


def main():
    parser = argparse.ArgumentParser(description="VoxDoc AI seller & ops CLI")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("init", help="Initialize the database")

    p = sub.add_parser("create-admin", help="Create an admin account")
    p.add_argument("email")
    p.add_argument("password")

    p = sub.add_parser("gen-keys", help="Mint sellable license keys")
    p.add_argument("tier", choices=["pro", "agency"])
    p.add_argument("--days", type=int, default=30, help="Validity in days after redemption")
    p.add_argument("--count", type=int, default=1, help="Number of keys to mint")

    p = sub.add_parser("list-keys", help="Show license keys")
    p.add_argument("--unredeemed", action="store_true", help="Only unredeemed keys")

    sub.add_parser("list-users", help="Show all accounts & usage")

    p = sub.add_parser("set-tier", help="Manually change a user's plan")
    p.add_argument("email")
    p.add_argument("tier", choices=["free", "pro", "agency"])
    p.add_argument("--days", type=int, default=30,
                   help="Validity in days (0 = never expires)")

    args = parser.parse_args()
    handlers = {
        "init": cmd_init,
        "create-admin": cmd_create_admin,
        "gen-keys": cmd_gen_keys,
        "list-keys": cmd_list_keys,
        "list-users": cmd_list_users,
        "set-tier": cmd_set_tier,
    }
    try:
        handlers[args.command](args)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
