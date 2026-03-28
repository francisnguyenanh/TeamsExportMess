"""Fix Phase 2 — replace lines 1192-1270 in app.py"""
import pathlib

app_path = pathlib.Path(__file__).parent / "app.py"
lines = app_path.read_text(encoding="utf-8").splitlines(keepends=True)

# Lines 1192-1270 (1-indexed) → indices 1191-1269
new_block = '''\
                # B2: set query /me/chats rồi Run
                if captured.get("source") != "graph_with_chat":
                    try:
                        log("📝 Thử query /me/chats…")
                        query_input = graph_page.locator(
                            'input[aria-label*="request URL"], '
                            'input[placeholder*="/me"], '
                            'input[type="url"], '
                            'input[role="combobox"]'
                        )
                        if query_input.count() > 0:
                            query_input.first.fill("https://graph.microsoft.com/v1.0/me/chats")
                            time.sleep(1)
                            run_btn = graph_page.locator(
                                'button:has-text("Run query"), button[aria-label*="Run"]'
                            )
                            if run_btn.count() > 0:
                                run_btn.first.click()
                                log("▶️ Clicked Run query for /me/chats")
                                time.sleep(5)
                    except Exception as e:
                        log(f"  ⚠️ Run query: {e}")

                # Step C: Auth Code + PKCE fallback
                if captured.get("source") != "graph_with_chat":
                    log("🔄 Auth Code + PKCE with Chat.Read + consent…")
                    try:
                        import hashlib
                        import secrets

                        code_verifier = secrets.token_urlsafe(64)
                        code_challenge = _b64_ge.urlsafe_b64encode(
                            hashlib.sha256(code_verifier.encode()).digest()
                        ).rstrip(b"=").decode()

                        GE_CLIENT_ID = "de8bc8b5-d9f9-48b1-a8ad-b748da725064"
                        REDIRECT_URI = "https://developer.microsoft.com/en-us/graph/graph-explorer"
                        SCOPES = "User.Read Chat.Read Chat.ReadBasic Team.ReadBasic.All Channel.ReadBasic.All ChannelMessage.Read.All openid profile offline_access"
                        auth_url = (
                            f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
                            f"?client_id={GE_CLIENT_ID}"
                            f"&response_type=code"
                            f"&redirect_uri={quote(REDIRECT_URI, safe='')}"
                            f"&scope={quote(SCOPES, safe='')}"
                            f"&code_challenge={code_challenge}"
                            f"&code_challenge_method=S256"
                            f"&prompt=consent"
                        )
                        log("📄 Consent page — bấm Accept nếu thấy popup…")
                        consent_page = context.new_page()
                        consent_page.goto(auth_url, timeout=60_000, wait_until="commit")

                        for wi in range(45):
                            time.sleep(2)
                            cur_url = consent_page.url
                            if "code=" in cur_url:
                                qs = parse_qs(urlparse(cur_url).query)
                                auth_code = qs.get("code", [""])[0]
                                if auth_code:
                                    log("📦 Got auth code, exchanging…")
                                    import requests as _req
                                    token_resp = _req.post(
                                        "https://login.microsoftonline.com/common/oauth2/v2.0/token",
                                        data={
                                            "client_id": GE_CLIENT_ID,
                                            "grant_type": "authorization_code",
                                            "code": auth_code,
                                            "redirect_uri": REDIRECT_URI,
                                            "code_verifier": code_verifier,
                                            "scope": SCOPES,
                                        },
                                        timeout=30,
                                    )
                                    if token_resp.status_code == 200:
                                        td = token_resp.json()
                                        captured["token"] = td["access_token"]
                                        captured["source"] = "graph_with_chat"
                                        log("✅ Graph token VỚI Chat.Read!")
                                    else:
                                        log(f"  ❌ Exchange: {token_resp.status_code} — {token_resp.text[:200]}")
                                break
                            if "error=" in cur_url:
                                frag = urlparse(cur_url).fragment
                                q_e = parse_qs(urlparse(cur_url).query)
                                f_e = parse_qs(frag) if frag else {}
                                err = q_e.get("error", f_e.get("error", [""]))[0]
                                desc = q_e.get("error_description", f_e.get("error_description", [""]))[0][:150]
                                log(f"  ❌ {err} — {desc}")
                                break
                            if wi % 5 == 4:
                                log(f"  ⏳ {(wi+1)*2}s — {cur_url[:80]}")
                        try:
                            consent_page.close()
                        except Exception:
                            pass
                    except Exception as e:
                        log(f"  ⚠️ Auth Code: {e}")

                # Final status
                if captured.get("source") == "graph_with_chat":
                    log("🎉 Graph token VỚI Chat.Read!")
                elif captured.get("source") == "graph_request":
                    _, scp, _ = _decode_token_scopes(captured.get("token", ""))
                    log(f"⚠️ Graph token nhưng KHÔNG có Chat.Read. Scopes: {scp[:120]}")
                    log("💡 Tip: Vào graph-explorer → Modify permissions → consent Chat.Read")

                try:
                    graph_page.close()
                except Exception:
                    pass
            except Exception as e:
                log(f"⚠️ Graph Explorer flow: {e}")
'''

new_lines = [l + '\n' for l in new_block.splitlines()]

# Replace lines 1192-1270 (indices 1191-1269)
result = lines[:1191] + new_lines + lines[1270:]
app_path.write_text("".join(result), encoding="utf-8")

print(f"Done! Old: {len(lines)} lines, New: {len(result)} lines")
print(f"Replaced lines 1192-1270 ({1270-1192+1} old lines) with {len(new_lines)} new lines")
