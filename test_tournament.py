#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
トーナメント表生成プログラムのテストスクリプト

仕様書の例に基づいてテストを実行し、結果を確認します。
"""

from tournament_generator import TournamentGenerator
import os


def test_tournament_generation():
    """トーナメント生成のテスト"""
    print("=== トーナメント表生成テスト ===\n")
    
    generator = TournamentGenerator()
    
    # テストケース1: 4人（仕様書の例1）
    print("テストケース1: 4人")
    participants_4 = generator.generate_participants(4)
    matches_4 = generator.generate_matches(participants_4)
    
    print(f"参加者: {participants_4}")
    print(f"対戦数: {len(matches_4)}試合")
    print("対戦カード:")
    for i, match in enumerate(matches_4, 1):
        print(f"  {i}. {match[0]} vs {match[1]}")
    
    # Excelファイル生成
    filename_4 = generator.generate_tournament(4)
    print(f"生成ファイル: {filename_4}")
    print(f"ファイル存在確認: {os.path.exists(filename_4)}")
    print()
    
    # テストケース2: 9人（仕様書の例2）
    print("テストケース2: 9人")
    participants_9 = generator.generate_participants(9)
    tables_9 = generator.split_participants(participants_9)
    
    print(f"参加者: {participants_9}")
    print(f"テーブル数: {len(tables_9)}個")
    
    for i, table in enumerate(tables_9):
        matches = generator.generate_matches(table)
        print(f"テーブル{chr(ord('A') + i)}: {len(table)}人, {len(matches)}試合")
        print(f"  参加者: {table}")
        print("  対戦カード:")
        for j, match in enumerate(matches, 1):
            print(f"    {j}. {match[0]} vs {match[1]}")
    
    # Excelファイル生成
    filename_9 = generator.generate_tournament(9)
    print(f"生成ファイル: {filename_9}")
    print(f"ファイル存在確認: {os.path.exists(filename_9)}")
    print()
    
    # テストケース3: 15人（3テーブル分割のテスト）
    print("テストケース3: 15人")
    participants_15 = generator.generate_participants(15)
    tables_15 = generator.split_participants(participants_15)
    
    print(f"参加者: {participants_15}")
    print(f"テーブル数: {len(tables_15)}個")
    
    for i, table in enumerate(tables_15):
        matches = generator.generate_matches(table)
        print(f"テーブル{chr(ord('A') + i)}: {len(table)}人, {len(matches)}試合")
        print(f"  参加者: {table}")
    
    # Excelファイル生成
    filename_15 = generator.generate_tournament(15)
    print(f"生成ファイル: {filename_15}")
    print(f"ファイル存在確認: {os.path.exists(filename_15)}")
    print()
    
    # テストケース4: 20人（4テーブル分割のテスト）
    print("テストケース4: 20人")
    participants_20 = generator.generate_participants(20)
    tables_20 = generator.split_participants(participants_20)
    
    print(f"参加者: {participants_20}")
    print(f"テーブル数: {len(tables_20)}個")
    
    for i, table in enumerate(tables_20):
        matches = generator.generate_matches(table)
        print(f"テーブル{chr(ord('A') + i)}: {len(table)}人, {len(matches)}試合")
        print(f"  参加者: {table}")
    
    # Excelファイル生成
    filename_20 = generator.generate_tournament(20)
    print(f"生成ファイル: {filename_20}")
    print(f"ファイル存在確認: {os.path.exists(filename_20)}")
    print()
    
    # テストケース5: 24人（上限テスト）
    print("テストケース5: 24人（上限）")
    participants_24 = generator.generate_participants(24)
    tables_24 = generator.split_participants(participants_24)
    
    print(f"参加者: {participants_24}")
    print(f"テーブル数: {len(tables_24)}個")
    
    for i, table in enumerate(tables_24):
        matches = generator.generate_matches(table)
        print(f"テーブル{chr(ord('A') + i)}: {len(table)}人, {len(matches)}試合")
        print(f"  参加者: {table}")
    
    # Excelファイル生成
    filename_24 = generator.generate_tournament(24)
    print(f"生成ファイル: {filename_24}")
    print(f"ファイル存在確認: {os.path.exists(filename_24)}")
    print()
    
    print("=== テスト完了 ===")
    print("生成されたファイル:")
    for filename in [filename_4, filename_9, filename_15, filename_20, filename_24]:
        if os.path.exists(filename):
            print(f"  ✓ {filename}")
        else:
            print(f"  ✗ {filename}")


def verify_combination_count():
    """組み合わせ数の検証"""
    print("\n=== 組み合わせ数検証 ===")
    
    generator = TournamentGenerator()
    
    test_cases = [3, 4, 5, 6, 7, 8, 9, 10, 15, 18, 20, 24]
    
    for n in test_cases:
        participants = generator.generate_participants(n)
        matches = generator.generate_matches(participants)
        expected = n * (n - 1) // 2  # nC2の計算
        
        print(f"{n}人: {len(matches)}試合 (期待値: {expected})")
        
        if len(matches) == expected:
            print("  ✓ 正しい組み合わせ数")
        else:
            print("  ✗ 組み合わせ数が間違っています")


if __name__ == "__main__":
    test_tournament_generation()
    verify_combination_count()
